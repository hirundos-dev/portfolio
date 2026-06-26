let observer = new IntersectionObserver(
  (elements) => {
    elements.forEach((el) => {
      if (el.intersectionRatio > 0.3) {
        el.target.classList.remove("sleep");
      } else {
        el.target.classList.add("sleep");
      }
    });
  },
  { threshold: [0, 0.5] }
);

document.querySelectorAll(".ids__snooze").forEach((el) => {
  observer.observe(el);
});
