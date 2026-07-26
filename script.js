const year = document.querySelector('#year');
const toggle = document.querySelector('.nav-toggle');
const navLinks = document.querySelector('#nav-links');

year.textContent = new Date().getFullYear();

toggle.addEventListener('click', () => {
  const isOpen = navLinks.classList.toggle('open');
  toggle.setAttribute('aria-expanded', String(isOpen));
});

navLinks.addEventListener('click', (event) => {
  if (event.target instanceof HTMLAnchorElement) {
    navLinks.classList.remove('open');
    toggle.setAttribute('aria-expanded', 'false');
  }
});
