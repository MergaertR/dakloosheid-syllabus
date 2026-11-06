const root = document.documentElement;
const themeBtn = document.getElementById("theme-toggle");

function applyTheme(theme) {
  root.dataset.theme = theme;
  themeBtn.textContent = theme === "dark" ? "Licht" : "Donker";
  localStorage.setItem("syllabus-theme", theme);
}
applyTheme(localStorage.getItem("syllabus-theme")
  || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"));
themeBtn.addEventListener("click", () => applyTheme(root.dataset.theme === "dark" ? "light" : "dark"));

function initToc() {
  const tocLinks = [...document.querySelectorAll(".toc a")];
  const observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      tocLinks.forEach((a) => a.classList.toggle("active", a.hash === "#" + entry.target.id));
    }
  }, { rootMargin: "-30% 0px -60% 0px" });
  document.querySelectorAll("main section[id]").forEach((s) => observer.observe(s));
}
if (document.querySelector("main section[id]")) initToc();

