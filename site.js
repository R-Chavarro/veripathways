const mobileBreakpoint = window.matchMedia("(max-width: 700px)");

document.querySelectorAll(".nav").forEach((nav) => {
    const toggle = nav.querySelector(".nav-toggle");
    const links = nav.querySelector(".nav-links");

    if (!toggle || !links) {
        return;
    }

    nav.classList.add("nav-enhanced");

    const closeMenu = () => {
        toggle.setAttribute("aria-expanded", "false");
        links.classList.remove("is-open");
    };

    toggle.addEventListener("click", () => {
        const isOpen = toggle.getAttribute("aria-expanded") === "true";
        toggle.setAttribute("aria-expanded", String(!isOpen));
        links.classList.toggle("is-open", !isOpen);
    });

    links.addEventListener("click", (event) => {
        if (event.target.closest("a")) {
            closeMenu();
        }
    });

    document.addEventListener("click", (event) => {
        if (!nav.contains(event.target)) {
            closeMenu();
        }
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && toggle.getAttribute("aria-expanded") === "true") {
            closeMenu();
            toggle.focus();
        }
    });

    mobileBreakpoint.addEventListener("change", closeMenu);
});

document.querySelectorAll("[data-current-year]").forEach((year) => {
    year.textContent = new Date().getFullYear();
});
