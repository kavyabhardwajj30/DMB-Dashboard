/**
 * Fast Client-Side Interactions for DMB Performance Dashboard
 * ==========================================================
 * Provides instant 0ms UI responsiveness:
 * 1. Instant Modal Dismissal (Close button, Backdrop click, Escape key)
 * 2. Instant Visual Feedback on clickables
 */

(function () {
    'use strict';

    function initInteractions() {
        // ----------------------------------------------------
        // 1. Navigation Tabs (if present)
        // ----------------------------------------------------
        var tabMpr = document.getElementById('nav-tab-mpr');
        var tabDmb = document.getElementById('nav-tab-dmb');

        function setActiveTab(activeName) {
            if (activeName === 'mpr') {
                if (tabMpr) tabMpr.classList.add('navigation-tab-active');
                if (tabDmb) tabDmb.classList.remove('navigation-tab-active');
            } else if (activeName === 'dmb') {
                if (tabDmb) tabDmb.classList.add('navigation-tab-active');
                if (tabMpr) tabMpr.classList.remove('navigation-tab-active');
            }
        }

        if (tabMpr && !tabMpr.hasAttribute('data-bound')) {
            tabMpr.setAttribute('data-bound', 'true');
            tabMpr.addEventListener('click', function (e) {
                e.preventDefault();
                setActiveTab('mpr');
                var mprTarget = document.getElementById('executive-insights-section') || document.getElementById('mpr-section');
                if (mprTarget) {
                    mprTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
                } else {
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }
            });
        }

        if (tabDmb && !tabDmb.hasAttribute('data-bound')) {
            tabDmb.setAttribute('data-bound', 'true');
            tabDmb.addEventListener('click', function (e) {
                e.preventDefault();
                setActiveTab('dmb');
                var dmbTarget = document.getElementById('dmb-section');
                if (dmbTarget) {
                    dmbTarget.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            });
        }

        // ----------------------------------------------------
        // 2. Instant Modal Close (Backdrop, Button, Escape)
        // ----------------------------------------------------
        function closeModalInstantly() {
            var modal = document.getElementById('continuous-red-modal');
            if (modal && !modal.classList.contains('continuous-red-modal-hidden')) {
                modal.classList.add('continuous-red-modal-hidden');
            }
        }

        var closeBtn = document.getElementById('close-continuous-red-modal');
        var backdrop = document.getElementById('continuous-red-modal-backdrop');

        if (closeBtn && !closeBtn.hasAttribute('data-bound')) {
            closeBtn.setAttribute('data-bound', 'true');
            closeBtn.addEventListener('click', closeModalInstantly);
        }
        if (backdrop && !backdrop.hasAttribute('data-bound')) {
            backdrop.setAttribute('data-bound', 'true');
            backdrop.addEventListener('click', closeModalInstantly);
        }

        // ----------------------------------------------------
        // 3. Instant Visual Feedback on RCA Card Click
        // ----------------------------------------------------
        var rcaCards = document.querySelectorAll('.continuous-red-panel-active');
        rcaCards.forEach(function (card) {
            if (!card.hasAttribute('data-click-bound')) {
                card.setAttribute('data-click-bound', 'true');
                card.addEventListener('click', function () {
                    card.classList.add('rca-card-opening');
                    setTimeout(function () {
                        card.classList.remove('rca-card-opening');
                    }, 800);
                });
            }
        });

        if (!window.__modalEscapeBound) {
            window.__modalEscapeBound = true;
            document.addEventListener('keydown', function (e) {
                if (e.key === 'Escape' || e.keyCode === 27) {
                    closeModalInstantly();
                }
            });
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initInteractions);
    } else {
        initInteractions();
    }

    // Re-bind after Dash reloads components
    var observer = new MutationObserver(function () {
        initInteractions();
    });
    observer.observe(document.body, { childList: true, subtree: true });
})();

