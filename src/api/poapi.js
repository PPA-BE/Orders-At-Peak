// src/api/poapi.js

import { msalInstance } from "../msal.js"; // MODIFIED: Import msalInstance

function toast(msg, isError = false) {
  const toastEl = document.getElementById("toast");
  if (toastEl) {
    toastEl.textContent = msg;
    toastEl.style.background = isError ? "#b91c1c" : "#16a34a";
    toastEl.classList.remove("hidden");
    setTimeout(() => toastEl.classList.add("hidden"), 2500);
  } else {
    isError ? console.error(msg) : console.log(msg);
  }
}

/**
 * NEW: A helper function to get the current user's identity from MSAL
 * and format it for API request headers. This ensures the backend
 * can properly log who performed the action.
 */
function getIdentityHeaders() {
  const headers = {};
  try {
    const acct = (msalInstance?.getActiveAccount?.() || (msalInstance?.getAllAccounts?.()[0])) || null;
    if (acct) {
      if (acct.username) headers["X-User-Email"] = acct.username;
      const name = acct.name || acct.idTokenClaims?.name;
      if (name) headers["X-User-Name"] = name;
    }
  } catch (e) {
    console.warn("Could not get user identity for API headers", e);
  }
  return headers;
}

export async function createPO(payload) {
  const res = await fetch("/api/po-create", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Failed to create PO in DB: ${res.status} ${body}`);
  }
  return res.json();
}

export async function saveEpicorPoNumber(id, epicorPoNumber) {
  const res = await fetch("/api/po-set-epicor", {
    method: "POST",
    // MODIFIED: Add identity headers
    headers: { "Content-Type": "application/json", ...getIdentityHeaders() },
    body: JSON.stringify({ id, epicorPoNumber }),
  });

  if (!res.ok) {
    const body = await res.text();
    toast(`Failed to save Epicor PO #: ${body}`, true);
    throw new Error(`API Error: ${res.status} ${body}`);
  }
  return res.json();
}

export async function updatePoStatus(id, status) {
  const res = await fetch(`/api/po-update-status`, {
    method: "POST",
    // MODIFIED: Add identity headers
    headers: { "Content-Type": "application/json", ...getIdentityHeaders() },
    body: JSON.stringify({ id, status }),
  });

  if (!res.ok) {
    const body = await res.text();
    toast(`Failed to update PO status: ${body}`, true);
    throw new Error(`API Error: ${res.status} ${body}`);
  }
  return res.json();
}

export async function markPoAsPaid(id) {
  const res = await fetch("/api/po-mark-paid", {
    method: "POST",
    // MODIFIED: Add identity headers
    headers: { "Content-Type": "application/json", ...getIdentityHeaders() },
    body: JSON.stringify({ id }),
  });

  if (!res.ok) {
    const body = await res.text();
    toast(`Failed to mark as paid: ${body}`, true);
    throw new Error(`API Error: ${res.status} ${body}`);
  }
  return res.json();
}

async function postJsonWithFallback(path1, path2, payload, extraHeaders = {}) {
  // Try primary path
  let res = await fetch(path1, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...extraHeaders },
    body: JSON.stringify(payload),
  });

  // If route not found / not allowed, retry the Netlify direct path
  if (!res.ok && (res.status === 404 || res.status === 405) && path2) {
    try {
      res = await fetch(path2, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...extraHeaders },
        body: JSON.stringify(payload),
      });
    } catch (_) { /* swallow, next block will throw */ }
  }
  if (!res.ok) {
    const body = await res.text();
    toast(`Failed to add payment: ${body}`, true);
    throw new Error(`API Error: ${res.status} ${body}`);
  }
  return res.json();
}

export async function addPoPayment({ id, amount, method, note }) {
  // REFACTORED: Use the centralized identity helper
  const headers = getIdentityHeaders();

  return postJsonWithFallback(
    "/api/po-add-payment",
    "/.netlify/functions/po-add-payment",
    { id, amount, method, note },
    headers
  );
}