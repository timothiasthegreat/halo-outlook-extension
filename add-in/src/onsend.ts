/**
 * OnMessageSend handler — intercepts outbound sends from Outlook
 * and journals them as ticket actions if the conversation is tracked.
 *
 * Registered via the Office.actions.associate("onMessageSend", handler)
 * pattern in the add-in manifest and taskpane initialization.
 */

import * as halo from "./halo";
import config from "./config";

/**
 * Called by Office.js before a message is sent.
 * Checks if the conversation is tracked and posts the sent email
 * to the Halo ticket as an "email_sent" action.
 *
 * Returns: { allowEvent: true } — always allow sending (even if
 * journaling fails, the email must go through).
 */
export async function onMessageSendHandler(): Promise<{ allowEvent: boolean }> {
  try {
    const conversationId = await halo.getCurrentConversationId();
    if (!conversationId) {
      return { allowEvent: true }; // Not tracked, let it send normally
    }

    // TODO: Check if this conversation is tracked (requires local state or Halo API).
    // For now, the watcher will pick up sent messages via Graph polling.
    // The add-in side requires either:
    //   a) A local state file (not available in browser context)
    //   b) A Halo API call to check the custom field (expensive per-send)
    //   c) A small config endpoint or localStorage cache of tracked conversations
    //
    // For Phase 3 MVP: rely on the watcher for outbound capture.
    // The OnSend handler is wired up but primarily serves as a future
    // enhancement for instant journaling (rather than waiting for the next poll cycle).

    // If we could determine the ticketId, we would post the action:
    // const config = getAddinConfig();
    // const subject = await halo.getCurrentSubject();
    // const body = await halo.getCurrentBody();
    // const internetMessageId = await halo.getCurrentInternetMessageId();
    //
    // await halo.createAction({
    //   ticket_id: ticketId,
    //   outcome_id: config.actions.emailSent,
    //   note: `Subject: ${subject}\n\n${stripHtml(body)}`,
    //   note_html: body,
    //   email_message_id: internetMessageId,
    //   sendemail: false, // Already being sent by Outlook
    //   hiddenfromuser: false,
    // });
  } catch {
    // Silently allow — never block an email send due to journal failure
  }

  return { allowEvent: true };
}

function stripHtml(html: string): string {
  if (typeof document !== "undefined") {
    const div = document.createElement("div");
    div.innerHTML = html;
    return div.textContent ?? div.innerText ?? "";
  }
  return html.replace(/<[^>]*>/g, "");
}