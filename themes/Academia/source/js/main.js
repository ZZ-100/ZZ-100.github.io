(function () {
    'use strict'

    document.addEventListener('DOMContentLoaded', function () {
        var header = document.querySelector('.header_wrap')
        var headerButton = document.querySelector('.menus_icon')
        var authorLinks = document.querySelector('.author-links')
        var socialButton = document.querySelector('.m-social-links')
        var nav = document.querySelector('.nav')
        var navButton = document.querySelector('.site-nav')
        var navWrap = document.querySelector('.nav-wrap')
        var topButton = document.querySelector('.cd-top')

        function setExpanded(button, expanded) {
            if (button) button.setAttribute('aria-expanded', expanded ? 'true' : 'false')
        }

        function closeHeader() {
            if (!header) return
            header.classList.remove('menus-open')
            header.classList.add('menus-close')
            setExpanded(headerButton, false)
        }

        function closeSocial() {
            if (!authorLinks) return
            authorLinks.classList.remove('is-open')
            authorLinks.classList.add('is-close')
            setExpanded(socialButton, false)
        }

        function closeNav() {
            if (!nav) return
            nav.classList.remove('nav-open')
            nav.classList.add('nav-close')
            setExpanded(navButton, false)
        }

        if (headerButton) {
            headerButton.addEventListener('click', function (event) {
                event.stopPropagation()
                var open = !header.classList.contains('menus-open')
                closeSocial()
                closeNav()
                header.classList.toggle('menus-open', open)
                header.classList.toggle('menus-close', !open)
                setExpanded(headerButton, open)
            })
        }

        if (socialButton) {
            socialButton.addEventListener('click', function (event) {
                event.stopPropagation()
                var open = !authorLinks.classList.contains('is-open')
                closeHeader()
                closeNav()
                authorLinks.classList.toggle('is-open', open)
                authorLinks.classList.toggle('is-close', !open)
                setExpanded(socialButton, open)
            })
        }

        if (navButton) {
            navButton.addEventListener('click', function (event) {
                event.stopPropagation()
                var open = !nav.classList.contains('nav-open')
                closeHeader()
                closeSocial()
                nav.classList.toggle('nav-open', open)
                nav.classList.toggle('nav-close', !open)
                setExpanded(navButton, open)
            })
        }

        document.addEventListener('click', function (event) {
            if (header && !event.target.closest('.header_wrap')) closeHeader()
            if (authorLinks && !event.target.closest('.author-links')) closeSocial()
            if (nav && !event.target.closest('.nav')) closeNav()
        })

        document.addEventListener('keydown', function (event) {
            if (event.key === 'Escape') {
                closeHeader()
                closeSocial()
                closeNav()
            }
        })

        function updateTopButton() {
            if (!navWrap) return
            navWrap.classList.toggle('is-visible', window.scrollY > 100)
        }

        window.addEventListener('scroll', updateTopButton, { passive: true })
        updateTopButton()

        if (topButton) {
            topButton.addEventListener('click', function () {
                window.scrollTo({ top: 0, behavior: 'smooth' })
            })
        }

        document.addEventListener('click', function (event) {
            var link = event.target.closest('a[href*="#"]')
            if (!link || link.getAttribute('href') === '#') return
            var url = new URL(link.href, window.location.href)
            if (url.pathname !== window.location.pathname || !url.hash) return
            var target = document.querySelector(url.hash)
            if (!target) return
            event.preventDefault()
            target.scrollIntoView({ behavior: 'smooth', block: 'start' })
            history.replaceState(null, '', url.hash)
        })
    })
})()
