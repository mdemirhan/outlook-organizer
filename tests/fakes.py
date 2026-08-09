from __future__ import annotations

from dataclasses import replace

from outlook_organizer.mail import FlagStatus, MailMessage


class FakeMailGateway:
    def __init__(self, messages: list[MailMessage]) -> None:
        self.messages = {message.outlook_id: message for message in messages}
        self.applied: list[dict] = []

    def latest_messages(self, folder_id, limit=20, body_limit=0):
        return list(self.messages.values())[:limit]

    def get_messages(self, outlook_ids, *, body_limit=0):
        return [replace(self.messages[value]) for value in outlook_ids if value in self.messages]

    def get_message(self, outlook_id, body_limit=0):
        return replace(self.messages[outlook_id])

    def apply_mail_states(self, updates):
        results = []
        for update in updates:
            self.apply_mail_state(
                update["outlook_id"],
                categories=update["categories"],
                flag_status=update["flag_status"],
                target_folder_id=update["target_folder_id"],
            )
            results.append({"outlook_id": update["outlook_id"], "status": "applied", "error": ""})
        return results

    def apply_mail_state(
        self,
        outlook_id,
        *,
        categories,
        flag_status: FlagStatus,
        target_folder_id=None,
    ):
        message = self.messages[outlook_id]
        self.applied.append(
            {
                "outlook_id": outlook_id,
                "categories": list(categories),
                "flag_status": flag_status,
                "target_folder_id": target_folder_id,
            }
        )
        message.categories = list(categories)
        message.flag_status = flag_status
        if target_folder_id is not None:
            message.folder_id = target_folder_id
