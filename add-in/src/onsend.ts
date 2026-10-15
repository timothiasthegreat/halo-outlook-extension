/**
 * OnMessageSend handler — intercepts outbound sends from Outlook
 * and journals them as ticket actions if the conversation is tracked.
 *
 * Registered via the Office.actions.associate("onMessageSend", handler)
 * pattern in the add-in manifest and taskpane initialization.
 */

import * as halo from "./halo";

/**
 * Quick check: do we have a usable Halo token in localStorage?
 * No network calls, no dialogs — pure synchronous localStorage read.
 */
function hasAuthToken(): boolean {
  try {
    const raw = localStorage.getItem("halo_outlook_token");
    if (!raw) return false;
    const token = JSON.parse(raw);
    return !!token.access_token && Date.now() < token.expires_at - 60000;
  } catch {
    return false;
  }
}

/**
 * Called by Office.js before a message is sent.
 *
 * Fast path: if the user is not authenticated (no stored token), allow
 * the send immediately — don't trigger the auth dialog or block the send.
 *
 * Otherwise, check if the conversation is tracked and post the sent email
 * to the Halo ticket as an "email_sent" action.
 *
 * Returns: { allowEvent: true } — always allow sending (even if
 * journaling fails, the email must go through).
 */
export async function onMessageSendHandler(): Promise<{ allowEvent: boolean }> {
  // ── Fast path: skip if not authenticated ──
  if (!hasAuthToken()) {
    return { allowEvent: true };
  }

  try {
    const conversationId = await halo.getCurrentConversationId();
    if (!conversationId) {
      return { allowEvent: true };
    }

    // Outbound mail is journaled by the watcher, not here. Finding the
    // ticket is now easy (halo.lookupConversation), but journaling at send
    // time would post every sent email twice: a compose item has no
    // internetMessageId until Exchange assigns one on send, so the action
    // posted here couldn't carry it, and the watcher's dedup (which matches
    // email_message_id on the ticket's actions) would post the sent copy
    // again on its next poll. The 2-second send timeout in launchevent.ts
    // also leaves little room for a lookup plus a large POST /Actions.

    // If we could determine the ticketId, we would post the action:
    // const cfg = (await import("./config")).getConfig();
    // const subject = await halo.getCurrentSubject();
    // const body = await halo.getCurrentBody();
    // const internetMessageId = await halo.getCurrentInternetMessageId();
    //
    // await halo.createAction({
    //   ticket_id: ticketId,
    //   outcome_id: cfg.actions.emailSent,
    //   note: `Subject: ${subject}\n\n${stripHtml(body)}`,
    //   note_html: body,
    //   email_message_id: internetMessageId,
    //   sendemail: false,
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