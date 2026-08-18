import { onMessageSendHandler } from "./onsend";

interface MailboxEvent {
  completed: (options: { allowEvent: boolean; errorMessage?: string }) => void;
}

// Track whether Office.js has been initialized for this runtime session
let initialized = false;

async function onMessageSendLaunch(event: MailboxEvent): Promise<void> {
  // Guard: ensure Office.js is ready before we touch Office.context.
  // Without this, OnMessageSend can fire before the runtime is initialized,
  // causing the handler to hang until SoftBlock's ~5s timeout.
  if (!initialized) {
    try {
      await Office.onReady();
      initialized = true;
    } catch {
      // Office.js failed to init — allow send immediately, don't block
      event.completed({ allowEvent: true });
      return;
    }
  }

  try {
    const result = await onMessageSendHandler();
    event.completed({ allowEvent: result.allowEvent !== false });
  } catch {
    // Never block send on runtime or service errors.
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendLaunch);