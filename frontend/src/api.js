const API_BASE = import.meta.env.VITE_API_BASE || '';

async function handleResponse(response) {
  if (!response.ok) {
    let message = `Request failed (${response.status})`;
    try {
      const body = await response.json();
      message = body.detail || message;
    } catch (_) {}
    throw new Error(message);
  }
  return response;
}

export async function parseExcel(file, headerRow = 0, formulas = 'values') {
  const form = new FormData();
  form.append('file', file);
  form.append('header_row', String(headerRow));
  form.append('formulas', formulas);
  const res = await handleResponse(await fetch(`${API_BASE}/api/parse-excel`, { method: 'POST', body: form }));
  return res.json();
}

export async function importWorkingSheet(file, headerRow = 0) {
  const form = new FormData();
  form.append('file', file);
  form.append('header_row', String(headerRow));
  const res = await handleResponse(await fetch(`${API_BASE}/api/import-working-sheet`, { method: 'POST', body: form }));
  return res.json();
}

export async function downloadExport(kind, payload) {
  const res = await handleResponse(await fetch(`${API_BASE}/api/export/${kind}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }));
  const blob = await res.blob();
  const disposition = res.headers.get('content-disposition') || '';
  const match = disposition.match(/filename="?([^";]+)"?/i);
  const filename = match?.[1] || `${payload.filename}.${kind === 'excel' ? 'xlsx' : 'pdf'}`;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
