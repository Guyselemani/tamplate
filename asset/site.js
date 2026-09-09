// This static template does not send messages to a server.
document.querySelector('#contact-form')?.addEventListener('submit', (event) => {
  event.preventDefault();
  document.querySelector('#form-status').textContent = 'Your message is ready. This demo is not connected to a delivery service; nothing has been sent.';
});
document.addEventListener('click', (event) => {
  document.querySelectorAll('.dropdown[open]').forEach((menu) => {
    if (!menu.contains(event.target)) menu.removeAttribute('open');
  });
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') document.querySelectorAll('.dropdown[open]').forEach((menu) => {
    menu.removeAttribute('open');
    menu.querySelector('summary').focus();
  });
});
