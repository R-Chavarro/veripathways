const mobileBreakpoint = window.matchMedia("(max-width: 700px)");
const precisePointer = window.matchMedia("(hover: hover) and (pointer: fine)");

document.querySelectorAll(".site-header").forEach((header) => {
    const menuToggle = header.querySelector(".nav-toggle");
    const hideAfter = 80;
    const revealZone = 48;
    let lastScrollY = window.scrollY;
    let scrollFrame;

    const menuIsOpen = () => menuToggle?.getAttribute("aria-expanded") === "true";
    const showHeader = () => header.classList.remove("header-hidden");
    const hideHeader = () => {
        if (window.scrollY > hideAfter && !menuIsOpen() && !header.contains(document.activeElement)) {
            header.classList.add("header-hidden");
        }
    };

    window.addEventListener("scroll", () => {
        if (scrollFrame) {
            return;
        }

        scrollFrame = window.requestAnimationFrame(() => {
            const currentScrollY = window.scrollY;

            if (currentScrollY <= hideAfter) {
                showHeader();
            } else if (!precisePointer.matches && currentScrollY < lastScrollY) {
                showHeader();
            } else {
                hideHeader();
            }

            lastScrollY = currentScrollY;
            scrollFrame = undefined;
        });
    }, { passive: true });

    document.addEventListener("pointermove", (event) => {
        if (!precisePointer.matches || window.scrollY <= hideAfter) {
            return;
        }

        if (event.clientY <= revealZone || header.contains(event.target)) {
            showHeader();
        } else {
            hideHeader();
        }
    }, { passive: true });

    header.addEventListener("focusin", showHeader);
});

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
