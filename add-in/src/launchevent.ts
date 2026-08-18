import { onMessageSendHandler } from "./onsend";

interface MailboxEvent {
  completed: (options: { allowEvent: boolean; errorMessage?: string }) => void;
}

/**
 * OnMessageSend launch event handler.
 *
 * Called by Outlook before a message is sent. Must complete quickly —
 * the PromptUser dialog appears if this handler takes too long.
 *
 * Strategy: always allow the send. Journaling to Halo is best-effort.
 * If anything goes wrong (auth missing, network error, timeout), we
 * bail out immediately and let the email go through.
 */
async function onMessageSendLaunch(event: MailboxEvent): Promise<void> {
  // Fire-and-forget: call the handler but don't block on it.
  // The PromptUser dialog shows briefly while we work, but the send
  // always proceeds regardless of handler outcome.
  onMessageSendHandler()
    .then((result) => {
      event.completed({ allowEvent: result.allowEvent !== false });
    })
    .catch(() => {
      event.completed({ allowEvent: true });
    });

  // Safety net: if the handler hangs (e.g. Office.js not ready),
  // allow the send after 2 seconds max.
  setTimeout(() => {
    try {
      event.completed({ allowEvent: true });
    } catch {
      // Already completed — no-op
    }
  }, 2000);
}

Office.actions.associate("onMessageSendHandler", onMessageSendLaunch);