/// <reference types="office-js" />
import { applyCidToSubject, getAllRecipients, getSubject, setSubject, showInfo } from './helpers';
import { getCidMatches } from './cidLookup';

interface EnsureResult {
  appliedCid?: string;
  multipleCids?: string[];
  error?: string;
}

async function ensureCidTag(): Promise<EnsureResult> {
  const item = Office.context.mailbox.item as Office.MessageCompose;
  const recipients = await getAllRecipients(item);
  if (!recipients.length) {
    return {};
  }

  const matches = await getCidMatches(recipients);
  if (!matches.length) {
    await showInfo(item, 'No CID found for current recipients. Consider adding this contact to the CID Register if appropriate.');
    return {};
  }

  const uniqueCids = Array.from(new Set(matches.map((m) => m.cid)));
  if (uniqueCids.length > 1) {
    await showInfo(item, `Multiple CIDs matched: ${uniqueCids.join(', ')}. Choose the correct CID before sending.`);
    return { multipleCids: uniqueCids };
  }

  const cid = uniqueCids[0];
  const subject = await getSubject(item);
  const nextSubject = applyCidToSubject(subject, cid);
  if (nextSubject !== subject) {
    await setSubject(item, nextSubject);
  }
  await showInfo(item, `CID applied: [${cid}]`);
  return { appliedCid: cid };
}

/** Event: fires on compose (new, reply, forward) */
export async function onNewMessageCompose(event: Office.AddinCommands.Event) {
  try {
    await ensureCidTag();
  } catch (err: any) {
    console.error('onNewMessageCompose error', err);
  } finally {
    event.completed();
  }
}

/** Event: fires when recipients change */
export async function onRecipientsChanged(event: Office.AddinCommands.Event) {
  try {
    await ensureCidTag();
  } catch (err: any) {
    console.error('onRecipientsChanged error', err);
  } finally {
    event.completed();
  }
}

/** Event: fires just before send. Use soft block if multiple CIDs found. */
export async function onMessageSend(event: Office.AddinCommands.Event) {
  try {
    const result = await ensureCidTag();
    if (result.multipleCids && result.multipleCids.length > 1) {
      const options: any = { allowEvent: false, errorMessage: `Multiple CIDs matched: ${result.multipleCids.join(', ')}. Remove ambiguous recipients or select one CID.` };
      event.completed(options);
      return;
    }
    event.completed({ allowEvent: true } as any);
  } catch (err: any) {
    console.error('onMessageSend error', err);
    const options: any = { allowEvent: false, errorMessage: 'CID validation failed. Please try again or check network/auth settings.' };
    event.completed(options);
  }
}

Office.actions.associate('onNewMessageCompose', onNewMessageCompose);
Office.actions.associate('onRecipientsChanged', onRecipientsChanged);
Office.actions.associate('onMessageSend', onMessageSend);
