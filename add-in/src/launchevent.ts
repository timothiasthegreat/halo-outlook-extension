import { onMessageSendHandler } from "./onsend";

interface MailboxEvent {
  completed: (options: { allowEvent: boolean; errorMessage?: string }) => void;
}

async function onMessageSendLaunch(event: MailboxEvent): Promise<void> {
  try {
    const result = await onMessageSendHandler();
    event.completed({ allowEvent: result.allowEvent !== false });
  } catch {
    // Never block send on runtime or service errors.
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendLaunch);
