let observer = new IntersectionObserver(
  (elements) => {
    elements.forEach((el) => {
      if (el.intersectionRatio > 0) {
        el.target.classList.remove("sleep");
        if (el.target.tagName === "VIDEO") el.target.play();
      } else {
        el.target.classList.add("sleep");
        if (el.target.tagName === "VIDEO") el.target.pause();
      }
    });
  },
  { threshold: [0, 0.5] }
);

document.querySelectorAll(".ids__snooze, video").forEach((el) => {
  observer.observe(el);
});
