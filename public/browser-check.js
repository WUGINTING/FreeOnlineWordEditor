/*
 * Old-browser notice (GOV-171). A classic script in old JavaScript on purpose: the application is
 * built for current browsers only (build target esnext), and an old one (Internet Explorer, Edge in
 * IE mode, a Chrome from before 2021) cannot even read it: the page would stay blank. This script
 * still runs there and says what to do instead.
 *
 * Two checks:
 * - right away: features every browser the application supports has: Array.prototype.at and
 *   Object.hasOwn (Chrome/Edge 93), the CSS selector :has() (Chrome/Edge 105, Firefox 121, Safari 15.4)
 *   and the inert attribute (Chrome/Edge 102, Firefox 112, Safari 15.5); the application's styles and
 *   dialogs use the last two, so the minimum is Chrome/Edge 105, Firefox 121, Safari 15.5;
 * - when the page has loaded: main.ts sets window.__appStarted; when it is still unset, the
 *   application could not start in this browser, whatever the first check said.
 * A separate file, not an inline script: the Content-Security-Policy allows scripts from 'self' only.
 */
(function () {
    var ID = 'dx-browser-notice';

    function supportsHas() {
        try {
            return !!(window.CSS && CSS.supports && CSS.supports('selector(:has(*))'));
        } catch (e) {
            return false;
        }
    }

    function supported() {
        return typeof Array.prototype.at === 'function'
            && typeof Object.hasOwn === 'function'
            && typeof window.Promise === 'function'
            && 'noModule' in document.createElement('script')
            && supportsHas()
            && 'inert' in document.createElement('div');
    }

    function show(started) {
        if (document.getElementById(ID) || !document.body) return;
        var ieMode = !!document.documentMode;
        var box = document.createElement('div');
        box.id = ID;
        box.setAttribute('role', 'alert');
        box.style.cssText = 'position:fixed;left:0;right:0;top:0;z-index:2147483647;padding:16px 20px;'
            + 'background:#fff4e5;color:#5f3700;border-bottom:2px solid #f0a020;'
            + 'font:16px/1.6 "Microsoft JhengHei","PingFang TC",sans-serif;';
        var title = document.createElement('strong');
        title.appendChild(document.createTextNode(started
            ? '您的瀏覽器版本較舊，部分功能可能無法正常使用。'
            : '這個瀏覽器無法開啟本系統。'));
        box.appendChild(title);
        var lines = [
            ieMode
                ? '目前是用 Internet Explorer（或 Edge 的「IE 模式」）開啟，請改用一般的 Microsoft Edge 或 Google Chrome。'
                : '請改用最新版的 Microsoft Edge 或 Google Chrome 開啟本網站。',
            '做法：複製下面的網址，貼到 Edge 或 Chrome 的網址列。電腦上沒有這些瀏覽器，或無法更新時，請洽單位的資訊人員。'
        ];
        for (var i = 0; i < lines.length; i++) {
            var p = document.createElement('div');
            p.appendChild(document.createTextNode(lines[i]));
            box.appendChild(p);
        }
        var url = document.createElement('input');
        url.type = 'text';
        url.readOnly = true;
        url.value = location.href;
        url.setAttribute('aria-label', '本網站網址');
        url.style.cssText = 'width:100%;max-width:640px;margin-top:6px;padding:4px 6px;font:inherit;box-sizing:border-box;';
        url.onfocus = function () { url.select(); };
        box.appendChild(url);
        if (started) {
            // The application runs (the first check may be too strict): the notice can be put away.
            var close = document.createElement('button');
            close.type = 'button';
            close.appendChild(document.createTextNode('我知道了，繼續使用'));
            close.style.cssText = 'margin-left:8px;padding:4px 12px;font:inherit;cursor:pointer;';
            close.onclick = function () { box.parentNode.removeChild(box); };
            box.appendChild(close);
        }
        document.body.appendChild(box);
    }

    function onReady(run) {
        if (document.readyState !== 'loading') run();
        else if (document.addEventListener) document.addEventListener('DOMContentLoaded', run);
        else window.attachEvent('onload', run);
    }

    function onLoad(run) {
        if (document.readyState === 'complete') run();
        else if (window.addEventListener) window.addEventListener('load', run);
        else window.attachEvent('onload', run);
    }

    if (!supported()) {
        // Shown once the application had its chance to start: it tells which of the two notices fits.
        onReady(function () { onLoad(function () { setTimeout(function () { show(!!window.__appStarted); }, 0); }); });
        return;
    }
    onLoad(function () {
        // Module scripts run before the load event: when main.ts has not run by then, it never will.
        setTimeout(function () { if (!window.__appStarted) show(false); }, 0);
    });
})();
