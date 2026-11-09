// Microsoft Graph calls for reading and updating the subject of a message.
const GRAPH = "https://graph.microsoft.com/v1.0";

// MAPI property PidTagSubject (0x0037, PT_UNICODE).
// Graph only allows PATCHing the "subject" property on drafts. Writing the MAPI property
// through singleValueExtendedProperties also works for received messages.
const PID_TAG_SUBJECT = "String 0x0037";

export class GraphError extends Error {
  constructor(status, code, message) {
    super(message);
    this.status = status;
    this.code = code;
  }
}

function messageUrl({ restId, mailbox }) {
  const owner = mailbox ? `users/${encodeURIComponent(mailbox)}` : "me";
  return `${GRAPH}/${owner}/messages/${encodeURIComponent(restId)}`;
}

async function send(token, url, init = {}) {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers || {}),
    },
  });
  if (!response.ok) {
    let code = "";
    let message = response.statusText;
    try {
      const body = await response.json();
      code = body?.error?.code || "";
      message = body?.error?.message || message;
    } catch {
      /* body is not JSON */
    }
    throw new GraphError(response.status, code, message);
  }
  return response.status === 204 ? null : response.json();
}

export async function getSubject(token, ref) {
  const data = await send(token, `${messageUrl(ref)}?$select=subject`);
  return data.subject ?? "";
}

export async function setSubject(token, ref, subject) {
  await send(token, messageUrl(ref), {
    method: "PATCH",
    body: JSON.stringify({
      singleValueExtendedProperties: [{ id: PID_TAG_SUBJECT, value: subject }],
    }),
  });
}
