/// <reference types="office-js" />
import { CONFIG } from './config';

export function extractExistingCid(subject: string): string | undefined {
  const match = subject.match(CONFIG.cidTagPattern);
  return match ? match[0] : undefined;
}

export function applyCidToSubject(subject: string, cidTag: string): string {
  const existing = extractExistingCid(subject);
  if (existing) return subject; // already tagged
  if (!subject || !subject.trim()) {
    return `[${cidTag}]`;
  }
  return `${subject.trim()} - [${cidTag}]`;
}

function wrapAsync<T>(fn: (callback: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

export async function getAllRecipients(item: Office.MessageCompose): Promise<string[]> {
  const to = await wrapAsync<Office.EmailAddressDetails[]>((cb) => item.to.getAsync({ asyncContext: null }, cb));
  const cc = await wrapAsync<Office.EmailAddressDetails[]>((cb) => item.cc.getAsync({ asyncContext: null }, cb));
  const bcc = item.bcc ? await wrapAsync<Office.EmailAddressDetails[]>((cb) => item.bcc!.getAsync({ asyncContext: null }, cb)) : [];
  const emails = [...to, ...cc, ...bcc]
    .map((r) => r?.emailAddress || r?.displayName)
    .filter(Boolean)
    .map((e) => e.trim().toLowerCase());
  return Array.from(new Set(emails));
}

export async function getSubject(item: Office.MessageCompose): Promise<string> {
  return wrapAsync<string>((cb) => item.subject.getAsync(cb));
}

export async function setSubject(item: Office.MessageCompose, subject: string): Promise<void> {
  return wrapAsync<void>((cb) => item.subject.setAsync(subject, cb));
}

export async function showInfo(item: Office.MessageCompose, message: string): Promise<void> {
  return wrapAsync<void>((cb) =>
    item.notificationMessages.replaceAsync(
      CONFIG.notifications.noMatch,
      { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message, icon: 'Icon.16x16', persistent: false },
      cb
    )
  );
}
