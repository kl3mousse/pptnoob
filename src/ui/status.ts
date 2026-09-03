export function setStatus(message: string): void {
  const status = document.getElementById("status");
  if (status) status.textContent = message;
}