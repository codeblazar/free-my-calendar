import logging
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo

import caldav
from caldav import DAVClient
from caldav.elements import dav
from dotenv import load_dotenv
from icalendar import Calendar, Event, vDatetime, vText
import win32com.client  # type: ignore
import locale


OL_DEFAULT_ITEMTYPE_APPOINTMENT = 1


LOGGER = logging.getLogger("calendarfree")
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=LOG_LEVEL, format="%(levelname)s %(message)s")


def get_env(name: str, default: Optional[str] = None, required: bool = False) -> str:
    value = os.getenv(name, default)
    if required and not value:
        sys.exit(f"Missing required environment variable: {name}")
    return value or ""


def ensure_delivery_method() -> None:
    delivery = os.getenv("DELIVERY_METHOD", "").lower()
    if delivery != "caldav":
        sys.exit("DELIVERY_METHOD must be set to 'caldav' in .env")


def resolve_timezone() -> ZoneInfo:
    tz_name = os.getenv("TIMEZONE", "Asia/Singapore")
    try:
        return ZoneInfo(tz_name)
    except Exception as exc:  # pragma: no cover - defensive
        raise SystemExit(f"Invalid TIMEZONE value '{tz_name}': {exc}") from exc


def _iter_com_collection(collection):
    item = collection.GetFirst()
    while item:
        yield item
        item = collection.GetNext()


def _iter_calendar_folders(namespace, *, max_depth: int = 5) -> Iterable[Tuple[str, "win32com.client.CDispatch"]]:
    """Yield (path, folder) for all calendar folders across all stores."""
    for store in _iter_com_collection(namespace.Folders):
        store_name = getattr(store, "Name", "") or "(Store)"
        queue: List[Tuple[object, str, int]] = [(store, store_name, 0)]
        while queue:
            folder, path, depth = queue.pop(0)
            try:
                default_type = getattr(folder, "DefaultItemType", None)
            except Exception:
                default_type = None
            if default_type == OL_DEFAULT_ITEMTYPE_APPOINTMENT:
                yield path, folder

            if depth >= max_depth:
                continue

            try:
                children = getattr(folder, "Folders", None)
            except Exception:
                children = None
            if not children:
                continue
            for child in _iter_com_collection(children):
                name = getattr(child, "Name", "?")
                queue.append((child, f"{path}/{name}", depth + 1))


def get_outlook_calendar_folder(
    calendar_name: str, store_name: Optional[str] = None
) -> "win32com.client.CDispatch":
    """Find a calendar folder by name, optionally constrained to a store."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    target = (calendar_name or "").strip().lower()
    store_target = (store_name or "").strip().lower()

    matches: List[Tuple[str, "win32com.client.CDispatch"]] = []
    for path, folder in _iter_calendar_folders(namespace):
        name = (getattr(folder, "Name", "") or "").strip().lower()
        if target and name != target:
            continue
        if store_target:
            store_root = path.split("/", 1)[0].strip().lower()
            if store_root != store_target and store_target not in path.lower():
                continue
        matches.append((path, folder))

    if not matches:
        hint = f" in store '{store_name}'" if store_name else ""
        LOGGER.warning("Calendar '%s'%s not found; using default Outlook calendar", calendar_name, hint)
        return namespace.GetDefaultFolder(9)  # olFolderCalendar

    # Prefer an exact store-root match if present
    if store_target:
        exact = [entry for entry in matches if entry[0].split("/", 1)[0].strip().lower() == store_target]
        if exact:
            matches = exact

    path, folder = matches[0]
    LOGGER.info("Using Outlook calendar '%s' (path: %s)", calendar_name or folder.Name, path)
    return folder


def _local_tzinfo():
    return datetime.now().astimezone().tzinfo


def _iter_items(items):
    item = items.GetFirst()
    while item:
        yield item
        item = items.GetNext()


def _restrict_time_local(dt: datetime) -> str:
    """Format a datetime string that Outlook Restrict understands, honoring locale."""
    local_dt = dt.astimezone(_local_tzinfo()) if dt.tzinfo else dt

    override = os.getenv("OUTLOOK_RESTRICT_FORMAT", "").strip()
    if override:
        return local_dt.strftime(override)

    try:
        locale.setlocale(locale.LC_TIME, "")
        date_fmt = locale.nl_langinfo(locale.D_FMT)
        time_fmt = locale.nl_langinfo(locale.T_FMT)
        return local_dt.strftime(f"{date_fmt} {time_fmt}")
    except Exception:
        pass

    try:
        return local_dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return local_dt.strftime("%m/%d/%Y %I:%M %p")


def _format_restriction(dt_start: datetime, dt_end: datetime) -> str:
    # Outlook expects M/D/YYYY hh:mm AM/PM
    fmt = "%m/%d/%Y %I:%M %p"
    return (
        f"[Start] >= '{dt_start.strftime(fmt)}' AND "
        f"[Start] <= '{dt_end.strftime(fmt)}'"
    )


def _ensure_tz(dt: datetime, tz: ZoneInfo) -> datetime:
    if dt.tzinfo is None:
        return dt.replace(tzinfo=tz)
    return dt.astimezone(tz)


def fetch_outlook_events(
    folder: "win32com.client.CDispatch",
    dt_start: datetime,
    dt_end: datetime,
    tz: ZoneInfo,
) -> List[dict]:
    results = []
    items = folder.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    restriction = (
        f"[Start] >= '{_restrict_time_local(dt_start)}' "
        f"AND [Start] <= '{_restrict_time_local(dt_end)}'"
    )
    try:
        candidates = items.Restrict(restriction)
    except Exception:
        LOGGER.warning("Outlook Restrict failed; using full item scan")
        candidates = items

    for item in _iter_items(candidates):
        try:
            start_utc = getattr(item, "StartUTC", None)
            end_utc = getattr(item, "EndUTC", None)
        except Exception:
            start_utc = None
            end_utc = None

        if start_utc and end_utc:
            start = start_utc.astimezone(tz)
            end = end_utc.astimezone(tz)
        else:
            start = _ensure_tz(getattr(item, "Start", dt_start), tz)
            end = _ensure_tz(getattr(item, "End", dt_start), tz)

        # Keep events that start within the window
        if not (dt_start <= start < dt_end):
            continue

        results.append(
            {
                "uid": getattr(item, "GlobalAppointmentID", None)
                or f"outlook-{getattr(item, 'EntryID', '')}-{int(start.timestamp())}",
                "subject": getattr(item, "Subject", "") or "Untitled",
                "location": getattr(item, "Location", ""),
                "body": (getattr(item, "Body", "") or "").strip(),
                "start": start,
                "end": end,
                "all_day": bool(getattr(item, "AllDayEvent", False)),
            }
        )

    return results


def build_icalendar_entry(event: dict) -> str:
    cal = Calendar()
    cal.add("prodid", "-//CalendarFree//EN")
    cal.add("version", "2.0")

    ev = Event()
    ev.add("uid", event["uid"])
    ev.add("summary", vText(event["subject"]))
    if event.get("location"):
        ev.add("location", vText(event["location"]))
    if event.get("body"):
        ev.add("description", vText(event["body"]))
    ev.add("dtstamp", datetime.now(timezone.utc))

    if event["all_day"]:
        ev.add("dtstart", event["start"].date())
        ev.add("dtend", event["end"].date())
    else:
        ev.add("dtstart", vDatetime(event["start"]))
        ev.add("dtend", vDatetime(event["end"]))

    cal.add_component(ev)
    return cal.to_ical().decode()


def find_icloud_calendar(
    principal: caldav.Principal, calendar_name: str
) -> Optional[caldav.Calendar]:
    for calendar in principal.calendars():
        props = calendar.get_properties([dav.DisplayName()])
        display_name = str(props.get(getattr(dav.DisplayName, "tag", "{DAV:}displayname"), ""))
        if not display_name:
            display_name = getattr(calendar, "name", "")
        if display_name.lower() == calendar_name.lower():
            return calendar
    return None


def connect_icloud_calendar(email: str, app_password: str, calendar_name: str) -> caldav.Calendar:
    client = DAVClient(
        url="https://caldav.icloud.com/",
        username=email,
        password=app_password,
    )
    principal = client.principal()
    calendar = find_icloud_calendar(principal, calendar_name)
    if not calendar:
        available = [
            str(
                cal.get_properties([dav.DisplayName()]).get(
                    getattr(dav.DisplayName, "tag", "{DAV:}displayname"), ""
                )
            )
            for cal in principal.calendars()
        ]
        raise SystemExit(
            f"iCloud calendar '{calendar_name}' not found. Available: {', '.join(available)}"
        )
    return calendar


def replace_icloud_calendar(calendar: caldav.Calendar, events_ics: Iterable[str]) -> None:
    existing = calendar.events()
    LOGGER.info("Deleting %d existing iCloud events...", len(existing))
    for event in existing:
        event.delete()

    events_list = list(events_ics)
    LOGGER.info("Uploading %d events to iCloud...", len(events_list))
    for ics in events_list:
        calendar.save_event(ics)


def main() -> None:
    load_dotenv(Path(".env"))
    ensure_delivery_method()

    tz = resolve_timezone()
    past_days = int(os.getenv("SYNC_DAYS_PAST", "3"))
    future_days = int(os.getenv("SYNC_DAYS_FUTURE", "60"))
    window_start = datetime.now(tz) - timedelta(days=past_days)
    window_end = datetime.now(tz) + timedelta(days=future_days)

    outlook_calendar_name = os.getenv("OUTLOOK_CALENDAR_NAME", "")
    outlook_store_name = os.getenv("OUTLOOK_STORE_NAME", "")
    folder = get_outlook_calendar_folder(outlook_calendar_name, outlook_store_name)
    outlook_events = fetch_outlook_events(folder, window_start, window_end, tz)
    LOGGER.info(
        "Fetched %d Outlook events between %s and %s",
        len(outlook_events),
        window_start.isoformat(),
        window_end.isoformat(),
    )

    if not outlook_events:
        LOGGER.warning("No Outlook events found in window; nothing to sync")

    icloud_email = get_env("ICLOUD_EMAIL", required=True)
    icloud_app_password = get_env("ICLOUD_APP_PASSWORD", required=True)
    icloud_calendar_name = get_env("ICLOUD_CALDAV_CALENDAR_NAME", required=True)

    calendar = connect_icloud_calendar(icloud_email, icloud_app_password, icloud_calendar_name)
    events_ics = [build_icalendar_entry(ev) for ev in outlook_events]
    replace_icloud_calendar(calendar, events_ics)
    LOGGER.info("Sync complete")


if __name__ == "__main__":
    main()
