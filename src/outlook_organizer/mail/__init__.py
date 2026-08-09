from outlook_organizer.mail.config import MailContext, load_mail_context
from outlook_organizer.mail.facts import FactBuilder
from outlook_organizer.mail.models import (
    Directness,
    DomainClass,
    FlagStatus,
    FolderCatalogConfig,
    FolderConfig,
    MailDefinitionsConfig,
    MailMessage,
    MessageFacts,
    OutlookFolder,
    Recipient,
)

__all__ = [
    "Directness",
    "DomainClass",
    "FactBuilder",
    "FlagStatus",
    "FolderCatalogConfig",
    "FolderConfig",
    "MailContext",
    "MailDefinitionsConfig",
    "MailMessage",
    "MessageFacts",
    "OutlookFolder",
    "Recipient",
    "load_mail_context",
]
