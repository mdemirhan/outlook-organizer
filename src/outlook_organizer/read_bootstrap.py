from __future__ import annotations

from outlook_organizer.brief import MailBriefService, load_brief_context
from outlook_organizer.calendar import CalendarService, load_calendar_config
from outlook_organizer.mail import load_mail_context
from outlook_organizer.mail.service import MailReadService
from outlook_organizer.outlook import OutlookReader


def mail_read_service() -> MailReadService:
    return MailReadService(load_mail_context(), OutlookReader())


def brief_service() -> MailBriefService:
    return MailBriefService(load_brief_context(), OutlookReader())


def calendar_service() -> CalendarService:
    return CalendarService(load_calendar_config(), OutlookReader())
