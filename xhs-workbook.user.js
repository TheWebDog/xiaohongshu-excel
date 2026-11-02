// ==UserScript==
// @name         小红书摸鱼助手：Excel/WPS表格伪装、图片视频浏览与下载
// @namespace    local.xhs.workbook
// @author       TheWebDog
// @license      MIT
// @homepageURL  https://github.com/TheWebDog/xiaohongshu-excel
// @supportURL   https://github.com/TheWebDog/xiaohongshu-excel/issues
// @version      0.4.5
// @description  小红书上班摸鱼：Excel/WPS工作簿表格伪装，图文视频浏览、图片视频下载、全部图片ZIP打包、评论续读、查看原文和快捷遮罩。
// @match        https://www.xiaohongshu.com/*
// @icon         data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDgiIGhlaWdodD0iNDgiIHZpZXdCb3g9IjAgMCA0OCA0OCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTkgMTcuNUM5IDE1LjAxNDcgMTEuMDE0NyAxMyAxMy41IDEzSDM5VjM5QzM5IDQwLjY1NjkgMzcuNjU2OSA0MiAzNiA0MkgxNUMxMS42ODYzIDQyIDkgMzkuMzEzNyA5IDM2VjE3LjVaIiBmaWxsPSJ1cmwoI3BhaW50MF9yYWRpYWxfNzY3ODBfMTQ1NTMpIi8+CjxwYXRoIGQ9Ik05IDE3LjVDOSAxNS4wMTQ3IDExLjAxNDcgMTMgMTMuNSAxM0gzOVYzOUMzOSA0MC42NTY5IDM3LjY1NjkgNDIgMzYgNDJIMTVDMTEuNjg2MyA0MiA5IDM5LjMxMzcgOSAzNlYxNy41WiIgZmlsbD0idXJsKCNwYWludDFfcmFkaWFsXzc2NzgwXzE0NTUzKSIgZmlsbC1vcGFjaXR5PSIwLjciLz4KPHBhdGggZD0iTTkgMjIuNDk5OUM5IDIwLjAxNDcgMTEuMDE0NyAxNy45OTk5IDEzLjUgMTcuOTk5OUgyN0MyNS4zNDMxIDE3Ljk5OTkgMjQgMTkuMzQzMSAyNCAyMC45OTk5VjI2Ljk5OTlDMjQgMjguNjU2OCAyMi42NTY5IDI5Ljk5OTkgMjEgMjkuOTk5OUgxNUMxMS42ODYzIDI5Ljk5OTkgOSAzMi42ODYyIDkgMzUuOTk5OVYyMi40OTk5WiIgZmlsbD0idXJsKCNwYWludDJfbGluZWFyXzc2NzgwXzE0NTUzKSIvPgo8cGF0aCBkPSJNOSAyMi40OTk5QzkgMjAuMDE0NyAxMS4wMTQ3IDE3Ljk5OTkgMTMuNSAxNy45OTk5SDI3QzI1LjM0MzEgMTcuOTk5OSAyNCAxOS4zNDMxIDI0IDIwLjk5OTlWMjYuOTk5OUMyNCAyOC42NTY4IDIyLjY1NjkgMjkuOTk5OSAyMSAyOS45OTk5SDE1QzExLjY4NjMgMjkuOTk5OSA5IDMyLjY4NjIgOSAzNS45OTk5VjIyLjQ5OTlaIiBmaWxsPSJ1cmwoI3BhaW50M19saW5lYXJfNzY3ODBfMTQ1NTMpIiBmaWxsLW9wYWNpdHk9IjAuMyIvPgo8cGF0aCBkPSJNOSAxMkM5IDguNjg2MjkgMTEuNjg2MyA2IDE1IDZIMjdWMThIMTVDMTEuNjg2MyAxOCA5IDIwLjY4NjMgOSAyNFYxMloiIGZpbGw9InVybCgjcGFpbnQ0X2xpbmVhcl83Njc4MF8xNDU1MykiLz4KPHBhdGggZD0iTTkgMTJDOSA4LjY4NjI5IDExLjY4NjMgNiAxNSA2SDI3VjE4SDE1QzExLjY4NjMgMTggOSAyMC42ODYzIDkgMjRWMTJaIiBmaWxsPSJ1cmwoI3BhaW50NV9yYWRpYWxfNzY3ODBfMTQ1NTMpIi8+CjxwYXRoIGQ9Ik05IDEyQzkgOC42ODYyOSAxMS42ODYzIDYgMTUgNkgyN1YxOEgxNUMxMS42ODYzIDE4IDkgMjAuNjg2MyA5IDI0VjEyWiIgZmlsbD0idXJsKCNwYWludDZfbGluZWFyXzc2NzgwXzE0NTUzKSIvPgo8cmVjdCB3aWR0aD0iMTUiIGhlaWdodD0iMTIiIHJ4PSIzIiB0cmFuc2Zvcm09Im1hdHJpeCgtMSAwIDAgMSAzOSA2KSIgZmlsbD0idXJsKCNwYWludDdfcmFkaWFsXzc2NzgwXzE0NTUzKSIvPgo8cmVjdCB4PSI0IiB5PSIyMyIgd2lkdGg9IjE2IiBoZWlnaHQ9IjE2IiByeD0iMy4yNSIgZmlsbD0idXJsKCNwYWludDhfcmFkaWFsXzc2NzgwXzE0NTUzKSIvPgo8cmVjdCB4PSI0IiB5PSIyMyIgd2lkdGg9IjE2IiBoZWlnaHQ9IjE2IiByeD0iMy4yNSIgZmlsbD0idXJsKCNwYWludDlfcmFkaWFsXzc2NzgwXzE0NTUzKSIgZmlsbC1vcGFjaXR5PSIwLjMiLz4KPHBhdGggZD0iTTE2LjIwNTYgMzUuNTcxNEgxMy43MzkzTDEyLjE5MDcgMzIuNjU3N0MxMi4xMzU1IDMyLjU1NTcgMTIuMDkzIDMyLjQ3NDkgMTIuMDYzMyAzMi40MTU0QzEyLjAzNzggMzIuMzUxNyAxMi4wMTAyIDMyLjI3OTQgMTEuOTgwNCAzMi4xOTg3SDExLjk1NDlDMTEuOTE2NyAzMi4zMDA3IDExLjg4MDYgMzIuMzgzNiAxMS44NDY2IDMyLjQ0NzNDMTEuODEyNiAzMi41MTExIDExLjc3MjIgMzIuNTg5NyAxMS43MjU1IDMyLjY4MzJMMTAuMTE5NiAzNS41NzE0SDcuNzkzNTJMMTAuNTg0OCAzMC45OTM2TDcuOTg0NyAyNi40Mjg2SDEwLjQxOTFMMTEuNzk1NiAyOS4wMjk5QzExLjg1MDggMjkuMTM2MiAxMS44OTc2IDI5LjIyOTcgMTEuOTM1OCAyOS4zMTA0QzExLjk3ODMgMjkuMzg2OSAxMi4wMjA4IDI5LjQ3ODMgMTIuMDYzMyAyOS41ODQ2SDEyLjA4ODhDMTIuMTQ4MiAyOS40NjEzIDEyLjE5NSAyOS4zNjM2IDEyLjIyOSAyOS4yOTEzQzEyLjI2NzIgMjkuMjE5MSAxMi4zMTgyIDI5LjEyMzQgMTIuMzgxOSAyOS4wMDQ0TDEzLjgwOTQgMjYuNDI4NkgxNi4xMjkxTDEzLjQ5MDggMzAuOTIzNUwxNi4yMDU2IDM1LjU3MTRaIiBmaWxsPSJ3aGl0ZSIvPgo8ZGVmcz4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDBfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKDM5IDQ0LjQxNjcpIHJvdGF0ZSgtMTMzLjU0OSkgc2NhbGUoNDEuMTIzNiAzMS45MDgyKSI+CjxzdG9wIG9mZnNldD0iMC4wNjQ3MjMzIiBzdG9wLWNvbG9yPSIjMzc5NTM5Ii8+CjxzdG9wIG9mZnNldD0iMC40MjIxNDkiIHN0b3AtY29sb3I9IiMyOTdDMkQiLz4KPHN0b3Agb2Zmc2V0PSIwLjcwMzQ0OCIgc3RvcC1jb2xvcj0iIzE1NTYxQyIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8cmFkaWFsR3JhZGllbnQgaWQ9InBhaW50MV9yYWRpYWxfNzY3ODBfMTQ1NTMiIGN4PSIwIiBjeT0iMCIgcj0iMSIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiIGdyYWRpZW50VHJhbnNmb3JtPSJ0cmFuc2xhdGUoMTguMTY2NyAyMy40NzIyKSByb3RhdGUoLTEzNS45NzEpIHNjYWxlKDE2LjY2MTEgMTIuODkwNikiPgo8c3RvcCBzdG9wLWNvbG9yPSIjMDczQjEwIi8+CjxzdG9wIG9mZnNldD0iMC45OTIyMzQiIHN0b3AtY29sb3I9IiMwODRBMTMiIHN0b3Atb3BhY2l0eT0iMCIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8bGluZWFyR3JhZGllbnQgaWQ9InBhaW50Ml9saW5lYXJfNzY3ODBfMTQ1NTMiIHgxPSI5IiB5MT0iMjYuOTk5OSIgeDI9IjIyLjc4NDgiIHkyPSIyNi45OTk5IiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CjxzdG9wIHN0b3AtY29sb3I9IiM1MkQxN0MiLz4KPHN0b3Agb2Zmc2V0PSIwLjMyODUxMiIgc3RvcC1jb2xvcj0iIzRBQTY0NyIvPgo8L2xpbmVhckdyYWRpZW50Pgo8bGluZWFyR3JhZGllbnQgaWQ9InBhaW50M19saW5lYXJfNzY3ODBfMTQ1NTMiIHgxPSIyMi41IiB5MT0iMTcuOTk5OSIgeDI9IjIyLjUiIHkyPSIzMC41MDI3IiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CjxzdG9wIHN0b3AtY29sb3I9IiMyOTg1MkYiLz4KPHN0b3Agb2Zmc2V0PSIwLjQ5OTk2OCIgc3RvcC1jb2xvcj0iIzRBQTY0NyIgc3RvcC1vcGFjaXR5PSIwIi8+CjwvbGluZWFyR3JhZGllbnQ+CjxsaW5lYXJHcmFkaWVudCBpZD0icGFpbnQ0X2xpbmVhcl83Njc4MF8xNDU1MyIgeDE9IjkuNjYzMDEiIHkxPSIyMC41ODk1IiB4Mj0iMjYuMjk0NCIgeTI9IjYuMTY0ODMiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KPHN0b3Agc3RvcC1jb2xvcj0iIzY2RDA1MiIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiM4NUU5NzIiLz4KPC9saW5lYXJHcmFkaWVudD4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDVfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKDI2LjM1MiAxMy42NjM0KSByb3RhdGUoLTE4MCkgc2NhbGUoOS4wMTg4MyAxOS4wOTQpIj4KPHN0b3Agb2Zmc2V0PSIwLjI5MjQ5NSIgc3RvcC1jb2xvcj0iIzRFQjQzQiIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiM3MkNDNjEiIHN0b3Atb3BhY2l0eT0iMCIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8bGluZWFyR3JhZGllbnQgaWQ9InBhaW50Nl9saW5lYXJfNzY3ODBfMTQ1NTMiIHgxPSIxNy4xMTQ3IiB5MT0iMTQuMzkyMyIgeDI9IjkiIHkyPSIxNC4zOTIzIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CjxzdG9wIG9mZnNldD0iMC4xODM2OTQiIHN0b3AtY29sb3I9IiNDMEUwNzUiIHN0b3Atb3BhY2l0eT0iMCIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNEMUVCOTUiLz4KPC9saW5lYXJHcmFkaWVudD4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDdfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKC0xLjkxMDk0IDEzLjY3OTQpIHJvdGF0ZSgtMzguOTY5Nykgc2NhbGUoMjEuNzUxIDIxLjY5MDQpIj4KPHN0b3Agb2Zmc2V0PSIwLjQzOTk4MyIgc3RvcC1jb2xvcj0iIzc5RTk2RCIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNEMEVCNzYiLz4KPC9yYWRpYWxHcmFkaWVudD4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDhfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKDQgMjMpIHJvdGF0ZSg0NSkgc2NhbGUoMjIuNjI3NCA2NC40MTQxKSI+CjxzdG9wIHN0b3AtY29sb3I9IiMyMEE4NUUiLz4KPHN0b3Agb2Zmc2V0PSIwLjk0Mzc1IiBzdG9wLWNvbG9yPSIjMDk0NDJBIi8+CjwvcmFkaWFsR3JhZGllbnQ+CjxyYWRpYWxHcmFkaWVudCBpZD0icGFpbnQ5X3JhZGlhbF83Njc4MF8xNDU1MyIgY3g9IjAiIGN5PSIwIiByPSIxIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIgZ3JhZGllbnRUcmFuc2Zvcm09InRyYW5zbGF0ZSgxMiAzMi42KSByb3RhdGUoOTApIHNjYWxlKDExLjIgMTIuOSkiPgo8c3RvcCBvZmZzZXQ9IjAuNTgwMzU3IiBzdG9wLWNvbG9yPSIjMzNBNjYyIiBzdG9wLW9wYWNpdHk9IjAiLz4KPHN0b3Agb2Zmc2V0PSIwLjk3MzgwNiIgc3RvcC1jb2xvcj0iIzk4RjBCMCIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8L2RlZnM+Cjwvc3ZnPgo=
// @grant        none
// @run-at       document-start
// ==/UserScript==

// Copyright (c) 2026 TheWebDog. Code: MIT.
// Embedded Microsoft artwork is excluded; see THIRD_PARTY_NOTICES.md.
(() => {
  "use strict";
  // Capture the original-page flag before the site's router normalizes the URL.
  if (new URL(location.href).searchParams.get('xhs_workbook_original')==='1') {
    window.__xhsWorkbook?.destroy?.();
    window.__xhsWorkbook={version:'0.4.5',show(){},destroy(){delete window.__xhsWorkbook;}};
    return;
  }
  if(window.__xhsWorkbook){
    if(window.__xhsWorkbook.version==='0.4.5'){window.__xhsWorkbook.show();return;}
    window.__xhsWorkbook.destroy?.();
  }
  if (!document.body) {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else boot();
  function boot() {
    if (window.__xhsWorkbook) return;
    const host = document.createElement("div");
    host.id = "xhs-workbook-host";
    host.dataset.version = "0.4.5";
    const shadow = host.attachShadow({ mode: "open" });
    const globalStyle = document.createElement("style");
    globalStyle.textContent =
      "body.xw-active > :not(#xhs-workbook-host):not(style):not(script):not(.login-modal),body.xw-active > :not(#xhs-workbook-host):not(style):not(script) *:not(.login-modal):not(.login-modal *){visibility:hidden!important;pointer-events:none!important}body.xw-active:not(.xw-panic) .login-modal.reds-modal-open{visibility:visible!important;pointer-events:auto!important;z-index:2147483647!important}body.xw-active.xw-panic > :not(#xhs-workbook-host):not(style):not(script){opacity:0!important;pointer-events:none!important}body.xw-active.xw-panic .login-modal,body.xw-active.xw-panic .login-modal *{visibility:hidden!important;pointer-events:none!important}";
    document.head.append(globalStyle);
    document.body.append(host);
    const $ = (s) => shadow.querySelector(s);
    const text = (e, s) => e?.querySelector(s)?.textContent?.trim() || "";
    const esc = (v) =>
      String(v ?? "").replace(
        /[&<>"']/g,
        (c) =>
          ({
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&#39;",
          })[c],
      );
    let enabled = true,
      panic = false,
      thumb = false,
      filter = "",
      sort = false,
      selected = -1,
      rows = [],
      signature = "",
      detailKey = "",
      detailSignature = "",
      detail = null,
      gallery = 0,
      mediaOpen = false,
      listTop = 0,
      pending = false,
      savedSearch = "",
      savedFilter = "";
    const seen = new Set();
    const cache = new Map();
    let context = location.pathname.includes("search_result")
      ? location.pathname + location.search
      : "/explore";
    shadow.innerHTML = `<style>
      :host{position:fixed;inset:0;z-index:2147483000;font:13px 'Microsoft YaHei','DengXian',sans-serif;color:#303030;line-height:1.45;--line:#d8d8d8;--green:#217346;display:block;pointer-events:none;color-scheme:light}
      *{box-sizing:border-box}button,input,select{font:inherit;color:inherit}button{cursor:pointer;border:1px solid transparent;background:transparent;border-radius:2px;padding:5px 10px}button:hover{background:#e0e9e2;border-color:#bccdc0}button:focus-visible,input:focus-visible{outline:2px solid var(--green);outline-offset:1px}button:disabled{opacity:.45;cursor:default}input{border:1px solid #c9c9c9;background:white;border-radius:2px;padding:5px 8px;min-width:0}a{color:#285e9b}
      .app,.restore{pointer-events:auto}.app{height:100dvh;display:flex;flex-direction:column;background:#fff}.title{height:33px;background:#217346;color:white;display:flex;align-items:center;padding:0 16px;gap:12px;flex-shrink:0}.logo img{width:20px;height:20px;display:block}.logo{display:grid;place-items:center;width:20px;height:20px;background:transparent;color:#217346;font-weight:bold}.filename{font-size:12px}.title .right{margin-left:auto;opacity:.85;font-size:12px}.tabs{height:34px;display:flex;align-items:center;gap:28px;padding:0 18px;background:#f5f5f5;border-bottom:1px solid var(--line);flex-shrink:0}.tabs .active{color:#217346;border-bottom:2px solid #217346;height:34px;display:grid;place-items:center}.ribbon{background:#f5f5f5;min-height:83px;display:flex;padding:8px 12px 4px;gap:0;border-bottom:1px solid #c8c8c8;flex-shrink:0}.group{padding:0 14px;border-right:1px solid #d5d5d5;display:flex;flex-direction:column;justify-content:space-between;gap:5px}.tools{display:flex;align-items:center;gap:5px;min-height:38px}.caption{text-align:center;color:#777;font-size:11px}.tool{display:flex;flex-direction:column;align-items:center;gap:2px;font-size:12px;padding:3px 8px;white-space:nowrap}.tool svg{width:22px;height:22px;fill:none;stroke:#527b62;stroke-width:1.4;stroke-linecap:round;stroke-linejoin:round}.fake-select{padding:4px 8px;border:1px solid #d2d2d2;background:white;white-space:nowrap;font-size:12px}.search{width:190px}.formula{height:36px;display:flex;align-items:center;gap:12px;padding:5px 8px;background:#f6f6f6;flex-shrink:0}.address{width:75px;background:white;border:1px solid #c9c9c9;text-align:center;font-size:12px;padding:3px}.fx{font:italic 18px Georgia;color:#888}.formula-value{flex:1;background:white;border:1px solid #c9c9c9;height:25px;padding:3px 9px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}.viewport{flex:1;min-height:0;overflow:auto;background:repeating-linear-gradient(0deg,transparent 0px,transparent 31px,#e7e7e7 31px,#e7e7e7 32px),repeating-linear-gradient(90deg,#fff 0px,#fff 119px,#e7e7e7 119px,#e7e7e7 120px);position:relative}table{border-collapse:separate;border-spacing:0;width:100%;table-layout:fixed;background:white}th,td{border-right:1px solid var(--line);border-bottom:1px solid var(--line);height:32px;padding:0 12px;text-align:left;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:normal}thead th{background:#ededed;color:#666;text-align:center;height:25px;position:sticky;top:0;z-index:2;font-size:12px}.num{width:42px;padding:0 5px;background:#f0f0f0;color:#777;text-align:center;font-variant-numeric:tabular-nums}.headers td{background:#f1f5f2;font-weight:600;color:#405c49;position:sticky;top:25px;z-index:1}.headers .num{background:#e5eae6}tbody tr.data:hover td:not(.num){background:#f4f8f5}tr.selected td:not(.num){background:#e9f2ec!important}tr.selected td.title-cell{outline:2px solid #217346;outline-offset:-2px}.title-cell{color:#245b92;cursor:pointer}.unread{display:inline-block;width:5px;height:5px;background:#b5b5b5;border-radius:50%;margin-right:9px;vertical-align:middle}.read .unread{background:transparent}.author-cell{color:#626262}.count{text-align:right;font-variant-numeric:tabular-nums}.attachment{color:#6b7570}.cover{width:44px;height:50px;object-fit:cover;vertical-align:middle;margin:5px 10px 5px 0}.thumb td{height:64px}.foot-tabs{height:33px;background:#f2f2f2;display:flex;gap:0;align-items:stretch;border-top:1px solid #c8c8c8;padding-left:16px;flex-shrink:0}.sheet{border-right:1px solid #ddd;padding:4px 20px;border-radius:0}.sheet.current{background:white;color:#217346;border-bottom:3px solid #217346;font-weight:600}.foot-tabs .hint{margin-left:auto;padding:7px 15px;color:#777;font-size:11px}.status{height:24px;background:#f3f3f3;border-top:1px solid #ddd;display:flex;align-items:center;gap:24px;padding:0 14px;color:#666;font-size:11px;flex-shrink:0}.status .right{margin-left:auto}.bottom{padding:12px 20px;background:#fff;color:#777;text-align:center;border-bottom:1px solid #ddd}.notice{padding:30px;background:white;color:#666;text-align:center}.detail-head{display:flex;align-items:center;gap:12px;padding:10px 16px;background:#f4f6f4;border-bottom:1px solid #ccc;position:sticky;top:0;z-index:3}.detail-head strong{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.detail-table td{white-space:normal;height:auto;padding:11px 16px;vertical-align:top;line-height:1.8}.detail-table .num{padding:10px 5px}.label{width:112px;color:#666;background:#f5f6f5}.prose{white-space:pre-wrap!important;word-break:break-word}.media{padding:14px;background:#fafafa;border:1px solid #ddd;text-align:center}.media img{display:block;max-width:100%;max-height:65vh;margin:10px auto;object-fit:contain}.media video{max-width:100%;max-height:65vh}.comment-table td{padding:8px 12px;white-space:pre-wrap;height:auto;vertical-align:top;line-height:1.7}.comment-table th{position:static}.restore{position:fixed;right:18px;bottom:18px;background:#217346;color:white;border:1px solid #fff;padding:9px 16px;box-shadow:0 2px 8px #0003}[hidden]{display:none!important}.pressed{background:#dce9df;border-color:#abc3b2}.mask td{color:#444}.mask .headers td{color:#405c49}@media(max-width:900px){.ribbon{overflow:auto}.group{padding:0 8px}.foot-tabs .hint{display:none}.tabs{gap:22px}table{min-width:760px}}

      .cell-link{padding:0;border:0!important;background:transparent!important;text-align:left;color:#245b92;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:middle;display:inline-block}.cell-link:hover{text-decoration:underline}.attachment button{color:#587363}.app.detail-view .ribbon{display:none}.detail-head{flex-wrap:wrap;gap:5px;min-height:43px}.detail-head .grow{flex:1;min-width:10px}.detail-head .position{color:#777;font-size:12px;padding:0 8px}.detail-head button{white-space:nowrap}.detail-head strong{max-width:180px}.comments-heading{padding:12px 16px;background:#edf2ee;border-block:1px solid #ccd5ce;display:flex;align-items:center;gap:14px}.comments-heading span{color:#777;font-size:12px}.comment-more,.list-more{padding:16px;background:white;text-align:center;border-bottom:1px solid #ddd}.comment-more button,.list-more button{border:1px solid #b8cbbb;background:#f2f7f3;min-width:160px;padding:8px 18px}.load-state{display:block;margin-top:7px;color:#777;font-size:12px}.reply-row td{background:#f9faf9}.reply-row .comment-body{padding-left:28px}.reply-toggle{display:block;margin-top:5px;color:#426f51;font-size:12px;padding:2px 0}.media{padding:0;overflow:hidden}.media-tools{height:42px;display:flex;justify-content:center;align-items:center;gap:16px;background:#f1f4f1;border-bottom:1px solid #ddd}.media-stage{height:clamp(220px,48vh,520px);display:grid;place-items:center;background:#fafafa}.media-stage img{margin:0;max-height:100%;max-width:100%;height:100%;object-fit:contain}.media-stage video{max-height:100%;max-width:100%}.prose{max-width:100%;line-height:1.85!important}.readable{max-width:1000px}.busy-note{padding:20px;background:white}.app.mask-view .foot-tabs .sheet,.app.mask-view .hint{visibility:hidden}.app.mask-view .ribbon input{visibility:hidden}
      .media-stage{position:relative;overflow:hidden}.media-stage img{position:absolute;inset:0;width:100%;height:100%;max-width:none;max-height:none;object-fit:contain}.media-stage.zoom{overflow:auto;display:block}.media-stage.zoom img{position:static;display:block;width:max(100%,1000px);height:auto;max-width:none;max-height:none}
      .video-tools{display:flex;align-items:center;gap:12px;padding:9px 12px;flex-wrap:wrap;background:#f1f4f1;border-bottom:1px solid #ddd}.video-tools button{border:1px solid #bfd0c3;min-width:70px;background:white}.video-tools input[type=range]{padding:0;accent-color:#217346}.video-progress{flex:1;min-width:100px}.video-volume{width:90px}.video-time{min-width:92px;font-variant-numeric:tabular-nums;color:#666;font-size:12px}.video-stage{background:#202423}.video-stage video{position:absolute!important;inset:0!important;display:block!important;width:100%!important;height:100%!important;max-width:none!important;max-height:none!important;object-fit:contain!important}.video-help{font-size:12px;color:#777;padding:7px 10px;background:#f7f8f7;text-align:left}
      .original-link{display:inline-block;padding:5px 10px;border:1px solid #bfd0c3;border-radius:2px;background:white;text-decoration:none;color:#365f43;white-space:nowrap}.original-link:hover{background:#e0e9e2}.download-tools{display:flex;align-items:center;gap:10px;padding:7px 10px;flex-wrap:wrap;border-bottom:1px solid #ddd;background:#f6f8f6;font-size:12px}.download-tools button{background:white;border:1px solid #c8d3ca}.download-status{color:#626c65}.download-help{padding:5px 10px;color:#777;background:#fafbfa;font-size:11px;text-align:left}
      .auth-notice{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:6px 14px;background:#fff8df;border-bottom:1px solid #e6d89d;color:#655629;font-size:12px;flex-shrink:0}.auth-notice button{border:1px solid #cabe8e;background:white;white-space:nowrap}.app.mask-view .auth-notice{display:none!important}
    </style>
    <main class="app">
      <div class="title"><span class="logo" aria-hidden="true"><img id="workbook-logo" alt=""></span><span class="filename">素材整理.xlsx</span><span class="right">自动保存已关闭　 ·　 本地视图</span></div>
      <div class="tabs"><span>文件</span><span class="active">开始</span><span>插入</span><span>页面布局</span><span>公式</span><span>数据</span><span>审阅</span><span>视图</span></div>
      <div class="ribbon">
        <div class="group"><div class="tools"><button class="tool" data-action="sync"><svg viewBox="0 0 24 24"><path d="M20 7V3l-3 3a8 8 0 1 0 2 11M20 3h-5"/></svg>同步记录</button></div><div class="caption">数据</div></div>
        <div class="group"><div class="tools"><span class="fake-select">等线　⌄</span><span class="fake-select">11　⌄</span><span style="font-size:16px">Ｂ　<i>I</i>　<u>U</u></span></div><div class="caption">字体</div></div>
        <div class="group"><div class="tools"><button class="tool" data-action="thumb"><svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16"/><path d="m4 17 5-6 5 5 3-3 4 4M15 8h1"/></svg>素材视图</button><button class="tool" data-action="sort"><svg viewBox="0 0 24 24"><path d="M5 4v16m-3-3 3 3 3-3M12 5h8M12 10h6M12 15h4M12 20h2"/></svg>点赞排序</button></div><div class="caption">排列与视图</div></div>
        <div class="group"><form class="tools" id="search-form"><input class="search" aria-label="搜索小红书" placeholder="搜索笔记，按 Enter"><button type="submit">查找</button></form><div class="caption">站内搜索</div></div>
        <div class="group"><div class="tools"><input id="filter" aria-label="筛选当前记录" placeholder="筛选当前记录" style="width:150px"></div><div class="caption">标题 / 作者</div></div>
        <div class="group"><div class="tools"><button class="tool" data-action="panic"><svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16"/><path d="M3 10h18M10 4v16"/></svg>普通报表</button><button class="tool" data-action="native"><svg viewBox="0 0 24 24"><path d="M9 5 2 12l7 7M2 12h14a5 5 0 0 1 5 5"/></svg>原始界面</button></div><div class="caption">Alt+Q / Alt+W</div></div>
      </div>
      <div class="auth-notice" id="auth-notice" role="status" hidden><span>当前未登录，搜索、更多评论等操作需要登录后使用。</span><button data-action="login">登录账户</button></div>
      <div class="formula"><span class="address">C2</span><span class="fx">ƒx</span><div class="formula-value">点击标题阅读笔记，点击附件直接看图</div></div>
      <div class="viewport" tabindex="0" aria-label="工作表"><section id="list"></section><section id="detail" hidden></section><section id="mask" hidden></section></div>
      <div class="foot-tabs"><button class="sheet" data-action="home">首页</button><button class="sheet current" data-action="list">推荐素材</button><button class="sheet" data-action="detail" id="detail-tab" hidden>记录详情</button><span class="hint">点击标题打开 · ↑↓ 选择 · Enter 打开 · Esc 返回 · Alt+Q 遮罩</span></div>
      <div class="status"><span id="state">就绪</span><span id="total"></span><span class="right">工作表视图　 ━━━━━●━━　100%</span></div>
    </main><button class="restore" data-action="native" hidden>返回工作簿 · Alt+W</button>`;

    const viewport = $(".viewport");
    let closing = false,
      closeKey = "",
      wantMedia = false,
      scrollMemory = 0,
      activeId = "",
      openSerial = 0;
    let commentsBusy = false,
      commentsFailed = false,
      commentsEnd = false,
      commentsBefore = 0,
      commentsTimer = 0,
      commentsGeneration = 0,
      lastLoadScroll = -1;
    let feedBusy = false,
      feedFailed = false,
      feedEnd = false,
      feedTimer = 0,
      feedBefore = 0;
    let zoomed = false;
    let maskScroll = 0,
      originalTitle = document.title,
      currentRoute = location.pathname + location.search;
    let lastBodySignature = "",
      lastCommentsSignature = "",
      scheduleTimer = 0;
    const timers = new Set();
    function later(fn, ms) {
      const t = setTimeout(() => {
        timers.delete(t);
        fn();
      }, ms);
      timers.add(t);
      return t;
    }

    const originalIcons = new Map();
    const workbookIcon = document.createElement('link');
    workbookIcon.rel = 'icon';
    workbookIcon.type = 'image/svg+xml';
    workbookIcon.sizes = 'any';
    // Original Microsoft Excel product artwork, downloaded from the official Excel page.
    // Source: https://www.microsoft.com/content/dam/microsoft/bade/images/icons/en-us/m365-app-icons-fy26/Excel-Icon-FY26.svg
    // Microsoft retains rights to its product artwork and trademarks.
    const workbookIconURL = "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDgiIGhlaWdodD0iNDgiIHZpZXdCb3g9IjAgMCA0OCA0OCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTkgMTcuNUM5IDE1LjAxNDcgMTEuMDE0NyAxMyAxMy41IDEzSDM5VjM5QzM5IDQwLjY1NjkgMzcuNjU2OSA0MiAzNiA0MkgxNUMxMS42ODYzIDQyIDkgMzkuMzEzNyA5IDM2VjE3LjVaIiBmaWxsPSJ1cmwoI3BhaW50MF9yYWRpYWxfNzY3ODBfMTQ1NTMpIi8+CjxwYXRoIGQ9Ik05IDE3LjVDOSAxNS4wMTQ3IDExLjAxNDcgMTMgMTMuNSAxM0gzOVYzOUMzOSA0MC42NTY5IDM3LjY1NjkgNDIgMzYgNDJIMTVDMTEuNjg2MyA0MiA5IDM5LjMxMzcgOSAzNlYxNy41WiIgZmlsbD0idXJsKCNwYWludDFfcmFkaWFsXzc2NzgwXzE0NTUzKSIgZmlsbC1vcGFjaXR5PSIwLjciLz4KPHBhdGggZD0iTTkgMjIuNDk5OUM5IDIwLjAxNDcgMTEuMDE0NyAxNy45OTk5IDEzLjUgMTcuOTk5OUgyN0MyNS4zNDMxIDE3Ljk5OTkgMjQgMTkuMzQzMSAyNCAyMC45OTk5VjI2Ljk5OTlDMjQgMjguNjU2OCAyMi42NTY5IDI5Ljk5OTkgMjEgMjkuOTk5OUgxNUMxMS42ODYzIDI5Ljk5OTkgOSAzMi42ODYyIDkgMzUuOTk5OVYyMi40OTk5WiIgZmlsbD0idXJsKCNwYWludDJfbGluZWFyXzc2NzgwXzE0NTUzKSIvPgo8cGF0aCBkPSJNOSAyMi40OTk5QzkgMjAuMDE0NyAxMS4wMTQ3IDE3Ljk5OTkgMTMuNSAxNy45OTk5SDI3QzI1LjM0MzEgMTcuOTk5OSAyNCAxOS4zNDMxIDI0IDIwLjk5OTlWMjYuOTk5OUMyNCAyOC42NTY4IDIyLjY1NjkgMjkuOTk5OSAyMSAyOS45OTk5SDE1QzExLjY4NjMgMjkuOTk5OSA5IDMyLjY4NjIgOSAzNS45OTk5VjIyLjQ5OTlaIiBmaWxsPSJ1cmwoI3BhaW50M19saW5lYXJfNzY3ODBfMTQ1NTMpIiBmaWxsLW9wYWNpdHk9IjAuMyIvPgo8cGF0aCBkPSJNOSAxMkM5IDguNjg2MjkgMTEuNjg2MyA2IDE1IDZIMjdWMThIMTVDMTEuNjg2MyAxOCA5IDIwLjY4NjMgOSAyNFYxMloiIGZpbGw9InVybCgjcGFpbnQ0X2xpbmVhcl83Njc4MF8xNDU1MykiLz4KPHBhdGggZD0iTTkgMTJDOSA4LjY4NjI5IDExLjY4NjMgNiAxNSA2SDI3VjE4SDE1QzExLjY4NjMgMTggOSAyMC42ODYzIDkgMjRWMTJaIiBmaWxsPSJ1cmwoI3BhaW50NV9yYWRpYWxfNzY3ODBfMTQ1NTMpIi8+CjxwYXRoIGQ9Ik05IDEyQzkgOC42ODYyOSAxMS42ODYzIDYgMTUgNkgyN1YxOEgxNUMxMS42ODYzIDE4IDkgMjAuNjg2MyA5IDI0VjEyWiIgZmlsbD0idXJsKCNwYWludDZfbGluZWFyXzc2NzgwXzE0NTUzKSIvPgo8cmVjdCB3aWR0aD0iMTUiIGhlaWdodD0iMTIiIHJ4PSIzIiB0cmFuc2Zvcm09Im1hdHJpeCgtMSAwIDAgMSAzOSA2KSIgZmlsbD0idXJsKCNwYWludDdfcmFkaWFsXzc2NzgwXzE0NTUzKSIvPgo8cmVjdCB4PSI0IiB5PSIyMyIgd2lkdGg9IjE2IiBoZWlnaHQ9IjE2IiByeD0iMy4yNSIgZmlsbD0idXJsKCNwYWludDhfcmFkaWFsXzc2NzgwXzE0NTUzKSIvPgo8cmVjdCB4PSI0IiB5PSIyMyIgd2lkdGg9IjE2IiBoZWlnaHQ9IjE2IiByeD0iMy4yNSIgZmlsbD0idXJsKCNwYWludDlfcmFkaWFsXzc2NzgwXzE0NTUzKSIgZmlsbC1vcGFjaXR5PSIwLjMiLz4KPHBhdGggZD0iTTE2LjIwNTYgMzUuNTcxNEgxMy43MzkzTDEyLjE5MDcgMzIuNjU3N0MxMi4xMzU1IDMyLjU1NTcgMTIuMDkzIDMyLjQ3NDkgMTIuMDYzMyAzMi40MTU0QzEyLjAzNzggMzIuMzUxNyAxMi4wMTAyIDMyLjI3OTQgMTEuOTgwNCAzMi4xOTg3SDExLjk1NDlDMTEuOTE2NyAzMi4zMDA3IDExLjg4MDYgMzIuMzgzNiAxMS44NDY2IDMyLjQ0NzNDMTEuODEyNiAzMi41MTExIDExLjc3MjIgMzIuNTg5NyAxMS43MjU1IDMyLjY4MzJMMTAuMTE5NiAzNS41NzE0SDcuNzkzNTJMMTAuNTg0OCAzMC45OTM2TDcuOTg0NyAyNi40Mjg2SDEwLjQxOTFMMTEuNzk1NiAyOS4wMjk5QzExLjg1MDggMjkuMTM2MiAxMS44OTc2IDI5LjIyOTcgMTEuOTM1OCAyOS4zMTA0QzExLjk3ODMgMjkuMzg2OSAxMi4wMjA4IDI5LjQ3ODMgMTIuMDYzMyAyOS41ODQ2SDEyLjA4ODhDMTIuMTQ4MiAyOS40NjEzIDEyLjE5NSAyOS4zNjM2IDEyLjIyOSAyOS4yOTEzQzEyLjI2NzIgMjkuMjE5MSAxMi4zMTgyIDI5LjEyMzQgMTIuMzgxOSAyOS4wMDQ0TDEzLjgwOTQgMjYuNDI4NkgxNi4xMjkxTDEzLjQ5MDggMzAuOTIzNUwxNi4yMDU2IDM1LjU3MTRaIiBmaWxsPSJ3aGl0ZSIvPgo8ZGVmcz4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDBfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKDM5IDQ0LjQxNjcpIHJvdGF0ZSgtMTMzLjU0OSkgc2NhbGUoNDEuMTIzNiAzMS45MDgyKSI+CjxzdG9wIG9mZnNldD0iMC4wNjQ3MjMzIiBzdG9wLWNvbG9yPSIjMzc5NTM5Ii8+CjxzdG9wIG9mZnNldD0iMC40MjIxNDkiIHN0b3AtY29sb3I9IiMyOTdDMkQiLz4KPHN0b3Agb2Zmc2V0PSIwLjcwMzQ0OCIgc3RvcC1jb2xvcj0iIzE1NTYxQyIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8cmFkaWFsR3JhZGllbnQgaWQ9InBhaW50MV9yYWRpYWxfNzY3ODBfMTQ1NTMiIGN4PSIwIiBjeT0iMCIgcj0iMSIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiIGdyYWRpZW50VHJhbnNmb3JtPSJ0cmFuc2xhdGUoMTguMTY2NyAyMy40NzIyKSByb3RhdGUoLTEzNS45NzEpIHNjYWxlKDE2LjY2MTEgMTIuODkwNikiPgo8c3RvcCBzdG9wLWNvbG9yPSIjMDczQjEwIi8+CjxzdG9wIG9mZnNldD0iMC45OTIyMzQiIHN0b3AtY29sb3I9IiMwODRBMTMiIHN0b3Atb3BhY2l0eT0iMCIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8bGluZWFyR3JhZGllbnQgaWQ9InBhaW50Ml9saW5lYXJfNzY3ODBfMTQ1NTMiIHgxPSI5IiB5MT0iMjYuOTk5OSIgeDI9IjIyLjc4NDgiIHkyPSIyNi45OTk5IiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CjxzdG9wIHN0b3AtY29sb3I9IiM1MkQxN0MiLz4KPHN0b3Agb2Zmc2V0PSIwLjMyODUxMiIgc3RvcC1jb2xvcj0iIzRBQTY0NyIvPgo8L2xpbmVhckdyYWRpZW50Pgo8bGluZWFyR3JhZGllbnQgaWQ9InBhaW50M19saW5lYXJfNzY3ODBfMTQ1NTMiIHgxPSIyMi41IiB5MT0iMTcuOTk5OSIgeDI9IjIyLjUiIHkyPSIzMC41MDI3IiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CjxzdG9wIHN0b3AtY29sb3I9IiMyOTg1MkYiLz4KPHN0b3Agb2Zmc2V0PSIwLjQ5OTk2OCIgc3RvcC1jb2xvcj0iIzRBQTY0NyIgc3RvcC1vcGFjaXR5PSIwIi8+CjwvbGluZWFyR3JhZGllbnQ+CjxsaW5lYXJHcmFkaWVudCBpZD0icGFpbnQ0X2xpbmVhcl83Njc4MF8xNDU1MyIgeDE9IjkuNjYzMDEiIHkxPSIyMC41ODk1IiB4Mj0iMjYuMjk0NCIgeTI9IjYuMTY0ODMiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KPHN0b3Agc3RvcC1jb2xvcj0iIzY2RDA1MiIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiM4NUU5NzIiLz4KPC9saW5lYXJHcmFkaWVudD4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDVfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKDI2LjM1MiAxMy42NjM0KSByb3RhdGUoLTE4MCkgc2NhbGUoOS4wMTg4MyAxOS4wOTQpIj4KPHN0b3Agb2Zmc2V0PSIwLjI5MjQ5NSIgc3RvcC1jb2xvcj0iIzRFQjQzQiIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiM3MkNDNjEiIHN0b3Atb3BhY2l0eT0iMCIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8bGluZWFyR3JhZGllbnQgaWQ9InBhaW50Nl9saW5lYXJfNzY3ODBfMTQ1NTMiIHgxPSIxNy4xMTQ3IiB5MT0iMTQuMzkyMyIgeDI9IjkiIHkyPSIxNC4zOTIzIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CjxzdG9wIG9mZnNldD0iMC4xODM2OTQiIHN0b3AtY29sb3I9IiNDMEUwNzUiIHN0b3Atb3BhY2l0eT0iMCIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNEMUVCOTUiLz4KPC9saW5lYXJHcmFkaWVudD4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDdfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKC0xLjkxMDk0IDEzLjY3OTQpIHJvdGF0ZSgtMzguOTY5Nykgc2NhbGUoMjEuNzUxIDIxLjY5MDQpIj4KPHN0b3Agb2Zmc2V0PSIwLjQzOTk4MyIgc3RvcC1jb2xvcj0iIzc5RTk2RCIvPgo8c3RvcCBvZmZzZXQ9IjEiIHN0b3AtY29sb3I9IiNEMEVCNzYiLz4KPC9yYWRpYWxHcmFkaWVudD4KPHJhZGlhbEdyYWRpZW50IGlkPSJwYWludDhfcmFkaWFsXzc2NzgwXzE0NTUzIiBjeD0iMCIgY3k9IjAiIHI9IjEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiBncmFkaWVudFRyYW5zZm9ybT0idHJhbnNsYXRlKDQgMjMpIHJvdGF0ZSg0NSkgc2NhbGUoMjIuNjI3NCA2NC40MTQxKSI+CjxzdG9wIHN0b3AtY29sb3I9IiMyMEE4NUUiLz4KPHN0b3Agb2Zmc2V0PSIwLjk0Mzc1IiBzdG9wLWNvbG9yPSIjMDk0NDJBIi8+CjwvcmFkaWFsR3JhZGllbnQ+CjxyYWRpYWxHcmFkaWVudCBpZD0icGFpbnQ5X3JhZGlhbF83Njc4MF8xNDU1MyIgY3g9IjAiIGN5PSIwIiByPSIxIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIgZ3JhZGllbnRUcmFuc2Zvcm09InRyYW5zbGF0ZSgxMiAzMi42KSByb3RhdGUoOTApIHNjYWxlKDExLjIgMTIuOSkiPgo8c3RvcCBvZmZzZXQ9IjAuNTgwMzU3IiBzdG9wLWNvbG9yPSIjMzNBNjYyIiBzdG9wLW9wYWNpdHk9IjAiLz4KPHN0b3Agb2Zmc2V0PSIwLjk3MzgwNiIgc3RvcC1jb2xvcj0iIzk4RjBCMCIvPgo8L3JhZGlhbEdyYWRpZW50Pgo8L2RlZnM+Cjwvc3ZnPgo=";
    $("#workbook-logo").src = workbookIconURL;
    function updateWorkbookIcon() {
      if (!enabled) return;
      document.head.querySelectorAll('link[rel]').forEach(link => {
        if (link !== workbookIcon && /(?:^|\s)(?:icon|apple-touch-icon(?:-precomposed)?|mask-icon)(?:\s|$)/i.test(link.rel)) {
          const key = link.rel.toLowerCase().split(/\s+/).filter(token => token !== 'shortcut').sort().join(' ') + '|' + (link.getAttribute('sizes') || '') + '|' + (link.getAttribute('media') || '');
          originalIcons.set(key, link);
          link.remove();
        }
      });
      if (workbookIcon.href !== workbookIconURL) workbookIcon.href = workbookIconURL;
      if (!workbookIcon.isConnected) document.head.append(workbookIcon);
    }
    function restoreOriginalIcons() {
      workbookIcon.remove();
      originalIcons.forEach(link => { if (!link.isConnected) document.head.append(link); });
      originalIcons.clear();
    }
    const iconObserver = new MutationObserver(updateWorkbookIcon);
    iconObserver.observe(document.head, {childList:true,subtree:true,attributes:true,attributeFilter:['rel','href']});
    let loginWasOpen = false, loginRequestPending = false, authMessage = '当前操作需要登录，请完成登录后重试';
    function loginDialog() {
      const dialog = document.querySelector('.login-modal.reds-modal-open');
      return dialog && dialog.getClientRects().length && getComputedStyle(dialog).display !== 'none' ? dialog : null;
    }
    function loginButton() {
      return [...document.querySelectorAll('.side-bar button.login-btn,button.login-btn')].find(button => button.getClientRects().length) || null;
    }
    function isLoggedOut() {
      return !!(loginDialog() || loginButton() || document.querySelector('input[placeholder*="登录探索"]'));
    }
    function requestLogin(reason = '此操作') {
      if (panic) return;
      pause();
      authMessage = reason === '账户登录' ? '请在弹窗中完成登录' : reason + '需要登录，请完成登录后重试';
      status(authMessage);
      if (loginDialog() || loginRequestPending) return;
      loginRequestPending = true;
      later(() => { loginRequestPending = false; }, 800);
      const button = loginButton();
      if (!loginDialog() && button) button.click();
      else if (!loginDialog() && enabled) native();
      later(sync, 0);
    }
    function syncAuthentication() {
      const dialog = loginDialog();
      $('.app').inert = !!dialog && !panic;
      const guest = isLoggedOut(), guestChanged = $('#auth-notice').hidden !== !guest;
      $('#auth-notice').hidden = !guest;
      if (guestChanged) { updateFeedState(); updateCommentState(); }
      if (dialog) loginRequestPending = false;
      if (dialog && !loginWasOpen && !panic) {
        pause();
        if (pending) { pending = false; closing = false; closeKey = ''; wantMedia = false; ++openSerial; detailNav(); }
        if (commentsBusy) { commentsBusy = false; commentsFailed = true; ++commentsGeneration; clearTimeout(commentsTimer); }
        if (feedBusy) { feedBusy = false; feedFailed = true; clearTimeout(feedTimer); }
        updateFeedState();
        updateCommentState();
        status(authMessage);
      }
      if (!dialog && loginWasOpen) authMessage = '当前操作需要登录，请完成登录后重试';
      loginWasOpen = !!dialog;
    }

    let videoLease = null,
      playbackRequested = false,
      playbackEpoch = 0,
      soundMuted = false,
      soundVolume = 1;
    function mediaElements() {
      return [
        ...new Set([
          ...document.querySelectorAll("video,audio"),
          ...shadow.querySelectorAll("video,audio"),
          ...(videoLease ? [videoLease.video] : []),
        ]),
      ];
    }
    function pause() {
      playbackRequested = false;
      playbackEpoch++;
      mediaElements().forEach((v) => {
        v.pause();
        v.muted = true;
      });
      updateVideoUI();
    }
    function canPlay(v) {
      return (
        enabled &&
        !panic &&
        mediaOpen &&
        !pending &&
        !closing &&
        !document.hidden &&
        videoLease?.video === v &&
        videoLease.key === detail?.key &&
        playbackRequested
      );
    }
    function guardMedia(event) {
      const v = event.target;
      if (!(v instanceof HTMLMediaElement)) return;
      if (!enabled && !panic) return;
      if (!canPlay(v)) {
        v.muted = true;
        v.pause();
      } else
        mediaElements().forEach((other) => {
          if (other !== v) {
            other.muted = true;
            other.pause();
          }
        });
    }
    function releaseVideo(nativeMode = false) {
      if (!videoLease) return;
      pause();
      const lease = videoLease;
      videoLease = null;
      const v = lease.video;
      lease.listeners.forEach(([name, fn]) => v.removeEventListener(name, fn));
      v.controls = lease.controls;
      v.loop = lease.loop;
      v.autoplay = false;
      v.style.cssText = lease.style;
      v.playbackRate = lease.rate;
      v.muted = nativeMode ? soundMuted : true;
      v.volume = soundVolume;
      if (lease.placeholder.parentNode?.isConnected)
        lease.placeholder.replaceWith(v);
      else {
        lease.placeholder.remove();
        v.remove();
      }
    }
    function mountVideo(stage, video) {
      if (videoLease?.video === video && video.parentNode === stage) {
        updateVideoUI();
        return;
      }
      releaseVideo();
      pause();
      const placeholder = document.createComment("workbook video position");
      const parent = video.parentNode;
      if (!parent) {
        status("视频播放器尚未准备好，请稍后重试");
        return;
      }
      parent.insertBefore(placeholder, video);
      videoLease = {
        video,
        key: detail.key,
        placeholder,
        style: video.style.cssText,
        controls: video.controls,
        loop: video.loop,
        rate: video.playbackRate,
        listeners: [],
      };
      video.controls = false;
      video.autoplay = false;
      video.removeAttribute("autoplay");
      video.loop = false;
      video.playsInline = true;
      video.muted = true;
      video.tabIndex = 0;
      video.setAttribute("aria-label", "笔记视频，空格播放或暂停");
      stage.append(video);
      for (const name of [
        "loadedmetadata",
        "durationchange",
        "timeupdate",
        "pause",
        "playing",
        "waiting",
        "ended",
        "error",
      ]) {
        const fn = () => {
          if (name === "pause" || name === "ended") playbackRequested = false;
          updateVideoUI();
        };
        video.addEventListener(name, fn);
        videoLease.listeners.push([name, fn]);
      }
      updateVideoUI();
    }
    function formatTime(value) {
      if (!Number.isFinite(value)) return "--:--";
      const seconds = Math.max(0, Math.floor(value));
      return (
        Math.floor(seconds / 60) + ":" + String(seconds % 60).padStart(2, "0")
      );
    }
    function updateVideoUI() {
      const v = videoLease?.video,
        b = $('[data-action="video-toggle"]');
      if (!v || !b) return;
      b.textContent = v.paused ? "播放" : "暂停";
      const range = $("#video-progress");
      if (range) {
        range.max = Number.isFinite(v.duration) ? v.duration : 0;
        range.disabled = !Number.isFinite(v.duration) || v.duration <= 0;
        if (shadow.activeElement !== range) range.value = v.currentTime || 0;
      }
      const time = $("#video-time");
      if (time)
        time.textContent =
          formatTime(v.currentTime) + " / " + formatTime(v.duration);
      const mute = $('[data-action="video-mute"]');
      if (mute)
        mute.textContent =
          soundMuted || soundVolume === 0 ? "开启声音" : "静音";
      const hint = $("#video-status");
      if (hint)
        hint.textContent = v.error
          ? "视频加载失败，可重新读取附件"
          : v.ended
            ? "播放结束，点击播放可重播"
            : v.paused
              ? "已暂停，点击播放后继续"
              : v.readyState < 3
                ? "缓冲中…"
                : "正在播放";
    }
    async function toggleVideo() {
      const v = videoLease?.video;
      if (!v) return;
      if (!v.paused) {
        pause();
        return;
      }
      if (
        !enabled ||
        panic ||
        !mediaOpen ||
        pending ||
        closing ||
        document.hidden
      )
        return;
      const epoch = ++playbackEpoch;
      playbackRequested = true;
      v.volume = soundVolume;
      v.muted = soundMuted || soundVolume === 0;
      if (v.ended) v.currentTime = 0;
      updateVideoUI();
      try {
        await v.play();
        if (epoch !== playbackEpoch || !canPlay(v)) {
          v.muted = true;
          v.pause();
        }
      } catch (error) {
        if (epoch !== playbackEpoch) return;
        playbackRequested = false;
        status(
          error.name === "NotAllowedError"
            ? "请点击视频播放按钮开始播放"
            : "视频未能播放，请重新读取附件",
        );
      }
      updateVideoUI();
    }
    function toggleSound() {
      soundMuted = !soundMuted;
      if (videoLease) {
        videoLease.video.muted = soundMuted || soundVolume === 0;
      }
      updateVideoUI();
    }
    function pauseHidden() {
      if (document.hidden && enabled) pause();
    }
    document.addEventListener("play", guardMedia, true);
    document.addEventListener("playing", guardMedia, true);
    shadow.addEventListener("play", guardMedia, true);
    shadow.addEventListener("playing", guardMedia, true);
    document.addEventListener("visibilitychange", pauseHidden);

    function status(msg) {
      $("#state").textContent = msg;
    }
    function show() {
      enabled = true;
      document.body.classList.add("xw-active");
      $(".app").hidden = false;
      $(".restore").hidden = true;
      sync();
    }
    function native() {
      if (enabled) {
        cancelDownload(true);
        pause();
        releaseVideo(true);
      }
      enabled = !enabled;
      document.body.classList.toggle("xw-active", enabled);
      $(".app").hidden = !enabled;
      $(".restore").hidden = enabled;
      if (enabled) {
        pause();
        sync();
      } else {
        restoreOriginalIcons();
        document.title = originalTitle;
        shadow.querySelectorAll("video").forEach((v) => v.pause());
      }
    }
    function readRows() {
      return [...document.querySelectorAll("section.note-item")].map(
        (el, i) => ({
          el,
          y: el.getBoundingClientRect().top + scrollY,
          id:
            el.getAttribute("data-note-id") ||
            el.querySelector("a.cover")?.pathname ||
            String(i),
          title: text(el, ".title") || "（无标题笔记）",
          author: text(el, ".author .name"),
          likes: text(el, ".like-wrapper .count"),
          cover: el.querySelector("a.cover img")?.src || "",
          video: !!el.querySelector('.play-icon,.video-icon,use[href*="play"]'),
          link: el.querySelector("a.title,a.cover"),
        }),
      );
    }
    function number(s) {
      return (parseFloat(s) || 0) * (/万|w/i.test(s) ? 10000 : 1);
    }
    function filtered() {
      const a = rows.filter((r) =>
        (r.title + " " + r.author).toLowerCase().includes(filter.toLowerCase()),
      );
      if (sort) a.sort((a, b) => number(b.likes) - number(a.likes));
      return a;
    }
    const col =
      '<colgroup><col style="width:42px"><col style="width:92px"><col style="width:84px"><col><col style="width:190px"><col style="width:96px"><col style="width:100px"></colgroup>';
    function renderList() {
      const top = viewport.scrollTop,
        a = filtered();
      selected = a.findIndex((r) => r.id === activeId);
      $("#total").textContent = a.length + " 条记录";
      $("#list").innerHTML =
        '<table aria-label="素材台账">' +
        col +
        '<thead><tr><th></th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th></tr></thead><tbody><tr class="headers"><td class="num">1</td><td>素材编号</td><td>类型</td><td>标题</td><td>作者</td><td>点赞数</td><td>附件</td></tr>' +
        a
          .map(
            (r, i) =>
              '<tr class="data ' +
              (thumb ? "thumb " : "") +
              (seen.has(r.id) ? "read " : "") +
              (selected === i ? "selected" : "") +
              '" data-row="' +
              i +
              '" tabindex="0"><td class="num">' +
              (i + 2) +
              "</td><td>" +
              String(i + 1).padStart(4, "0") +
              "</td><td>" +
              (r.video ? "视频" : "图文") +
              '</td><td class="title-cell"><span class="unread"></span>' +
              (thumb && r.cover
                ? '<img class="cover" loading="lazy" src="' +
                  esc(r.cover) +
                  '" alt="笔记封面">'
                : "") +
              '<button class="cell-link" data-action="open" data-index="' +
              i +
              '" title="' +
              esc(r.title) +
              '">' +
              esc(r.title) +
              '</button></td><td class="author-cell">' +
              esc(r.author) +
              '</td><td class="count">' +
              (esc(r.likes) || "—") +
              '</td><td class="attachment"><button class="cell-link" data-action="attachment" data-index="' +
              i +
              '">' +
              (r.video ? "查看视频" : "查看图片") +
              "</button></td></tr>",
          )
          .join("") +
        "</tbody></table>" +
        (!a.length
          ? '<div class="notice">' +
            (rows.length
              ? "没有符合筛选条件的记录"
              : "等待小红书加载记录，可切回原界面检查") +
            "</div>"
          : "") +
        '<div class="list-more"><button data-action="more">继续加载记录</button><span class="load-state"></span></div>';
      $('.sheet[data-action="list"]').textContent = context.includes(
        "search_result",
      )
        ? "搜索结果"
        : "推荐素材";
      viewport.scrollTop = top;
      updateFeedState();
    }
    function updateFeedState() {
      const b = $(".list-more button"),
        s = $(".list-more .load-state");
      if (!b) return;
      b.disabled = feedBusy || feedEnd;
      b.textContent = feedBusy
        ? "正在加载…"
        : feedEnd
          ? "已显示全部记录"
          : feedFailed
            ? "重试加载记录"
            : "继续加载记录";
      s.textContent = feedFailed
        ? isLoggedOut() ? "当前未登录，请登录后重试加载记录。" : "暂未收到新记录，可重试或切回原界面检查。"
        : "已读取 " + rows.length + " 条；滚动到底可继续加载";
    }
    function select(i) {
      const r = filtered()[i];
      if (!r) return;
      selected = i;
      activeId = r.id;
      shadow
        .querySelectorAll("tr.data")
        .forEach((e) => e.classList.toggle("selected", +e.dataset.row === i));
      $(".address").textContent = "C" + (i + 2);
      $(".formula-value").textContent = r.title;
    }
    function open(i, attachments = false) {
      if (pending) return;
      const r = filtered()[i];
      if (!r?.link) return;
      const changing = !!detail;
      select(i);
      seen.add(r.id);
      if (!changing) listTop = viewport.scrollTop;
      wantMedia = attachments;
      mediaOpen = attachments;
      gallery = 0;
      pending = true;
      if (detail) detailNav();
      const serial = ++openSerial;
      status("正在读取记录…");
      const clickLive = () => {
        const live = [...document.querySelectorAll("section.note-item")].find(
          (e) => e.getAttribute("data-note-id") === r.id,
        );
        const link = live?.querySelector("a.title,a.cover");
        if (!link) return false;
        link.click();
        return true;
      };
      const navigate = () => {
        if (serial !== openSerial) return;
        if (clickLive()) return;
        window.scrollTo(0, Math.max(0, r.y - 180));
        let attempts = 0;
        const retry = () => {
          if (serial !== openSerial) return;
          if (!clickLive() && ++attempts < 25) later(retry, 140);
          else if (attempts >= 25) {
            pending = false;
            status("记录未恢复，请同步后重试");
          }
        };
        later(retry, 150);
      };
      if (changing) {
        closing = true;
        closeKey = detail.key;
        closeNative();
        let n = 0;
        const waitClosed = () => {
          if (serial !== openSerial) return;
          if (!document.querySelector("#noteContainer")) navigate();
          else if (++n < 35) later(waitClosed, 100);
          else {
            pending = false;
            closing = false;
            status("上一条详情未关闭，请返回台账后重试");
          }
        };
        later(waitClosed, 100);
      } else navigate();
      later(() => {
        if (serial === openSerial && pending) {
          pending = false;
          status(isLoggedOut() ? "当前未登录，请点击“登录账户”后重试" : "详情未加载，可切回原界面查看提示");
        }
      }, 10000);
    }
    function closeNative() {
      cancelDownload(true);
      pause();
      releaseVideo();
      const close =
        document.querySelector(".note-detail-mask .close") ||
        document.querySelector(".note-detail-mask .close-circle") ||
        document.querySelector(".note-detail-mask button.close-icon");
      if (close) close.click();
      else if (detail) history.back();
    }

    function pageNote(key){
      const unref=value=>value?._value??value?.value??value;
      const state=window.__INITIAL_STATE__?.note;
      const map=unref(state?.noteDetailMap);
      const entry=unref(map?.[key]);
      return unref(entry?.note)||null;
    }
    function resourceURL(raw){
      if(typeof raw!=='string'||!raw)return '';
      try{const url=new URL(raw,location.href);if(!['http:','https:'].includes(url.protocol))return '';
        if(!/(^|\.)(xhscdn\.com|xiaohongshu\.com)$/.test(url.hostname))return '';
        if(url.protocol==='http:')url.protocol='https:';return url.href;
      }catch{return '';}
    }
    function noteVideoFiles(note,nativeVideo){
      if(note?.video?.media?.video?.drmType)return [];
      let stream=note?.video?.media?.stream;
      if(!stream&&note?.video?.mediaV2){try{stream=JSON.parse(note.video.mediaV2).stream;}catch{}}
      const hasAudio=item=>!!(item.audioChannels||item.audio_channels||item.audioCodec||item.audio_codec);
      const candidates=Object.values(stream||{}).flat().filter(item=>item&&String(item.format).toLowerCase()==='mp4')
        .sort((a,b)=>Number(hasAudio(b))-Number(hasAudio(a))||(b.width||0)*(b.height||0)-(a.width||0)*(a.height||0))
        .flatMap(item=>[item.masterUrl||item.master_url,...(item.backupUrls||item.backup_urls||[])])
        .map(resourceURL).filter(Boolean);
      return [...new Set([...candidates,...[nativeVideo?.currentSrc,nativeVideo?.src,...(nativeVideo?.querySelectorAll('source')||[])].map(value=>resourceURL(typeof value==='string'?value:value?.src)).filter(Boolean)])];
    }
    function originalNoteURL(key,note){
      let url;try{const current=new URL(location.href);url=current.pathname.includes(key)?current:new URL(cache.get(key)?.link?.href||location.href);}catch{return '';}
      if(!url.pathname.includes(key)){
        url=new URL('/explore/'+encodeURIComponent(key),location.origin);
        if(note?.xsecToken)url.searchParams.set('xsec_token',note.xsecToken);
        url.searchParams.set('xsec_source','pc_feed');
      }
      if(!url.searchParams.get('xsec_token')&&note?.xsecToken)url.searchParams.set('xsec_token',note.xsecToken);
      url.searchParams.set('xhs_workbook_original','1');return url.href;
    }
    let downloadJob=null,downloadMessage='',downloadKey='',downloadReady=null;const downloadURLs=new Set();
    const safeFilename=value=>String(value||'小红书笔记').replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').replace(/[. ]+$/g,'')||'小红书笔记';
    const resourceBasename=note=>Array.from(safeFilename(note.title)).slice(0,60).join('')+'_'+safeFilename(note.key);
    function downloadToolbar(){
      const box=$('#media-slot .media');if(!box||!detail)return;
      let tools=box.querySelector('.download-tools');if(!tools){tools=document.createElement('div');tools.className='download-tools';box.prepend(tools);}
      tools.innerHTML=detail.video?'<button data-action="download-video">下载视频 MP4</button>':'<button data-action="download-image">下载当前图片</button><button data-action="download-all">全部图片 ZIP（'+detail.images.length+'）</button>';
      tools.innerHTML+='<button data-action="download-cancel" hidden>取消下载</button><span class="download-status" role="status"></span>';
      updateDownloadUI();
    }
    function updateDownloadUI(){
      const tools=$('.download-tools');if(!tools)return;const active=!!(downloadJob&&detail&&downloadJob.key===detail.key);
      tools.querySelectorAll('[data-action="download-image"],[data-action="download-all"],[data-action="download-video"]').forEach(b=>b.disabled=!!downloadJob);
      tools.querySelector('[data-action="download-cancel"]').hidden=!active;
      tools.querySelector('.download-status').textContent=downloadKey===detail?.key?downloadMessage:'';
      tools.querySelector('.download-save')?.remove();if(downloadReady&&detail&&downloadReady.key===detail.key){const a=document.createElement('a');a.className='download-save original-link';a.href=downloadReady.url;a.download=downloadReady.name;a.textContent='保存文件';tools.append(a);}
    }
    function cancelDownload(silent=false){if(downloadReady){URL.revokeObjectURL(downloadReady.url);downloadURLs.delete(downloadReady.url);downloadReady=null;}if(silent){downloadKey='';downloadMessage='';}if(!downloadJob){updateDownloadUI();return;}const key=downloadJob.key;downloadJob.controller.abort();downloadJob=null;downloadKey=silent?'':key;downloadMessage=silent?'':'已取消下载';updateDownloadUI();}
    function checkDownload(job){if(job.controller.signal.aborted||downloadJob!==job||detail?.key!==job.key)throw new DOMException('Download cancelled','AbortError');}
    function downloadProgress(job,message){checkDownload(job);downloadKey=job.key;downloadMessage=message;updateDownloadUI();}
    async function mp4HasTracks(blob,job){
      const handlers=new Set();let mediaData=false;
      async function boxes(start,end,depth){
        if(depth>4)return;
        for(let offset=start;offset+8<=end;){
          checkDownload(job);
          const head=new DataView(await blob.slice(offset,offset+16).arrayBuffer());
          let size=head.getUint32(0),header=8;
          const type=String.fromCharCode(...new Uint8Array(head.buffer,4,4));
          if(size===1){if(head.byteLength<16)return;size=head.getUint32(8)*4294967296+head.getUint32(12);header=16;}
          if(size===0)size=end-offset;
          if(!Number.isSafeInteger(size)||size<header||offset+size>end)return;
          if(type==='mdat'&&depth===0&&size>header)mediaData=true;
          if(type==='hdlr'&&size>=header+12){const value=new Uint8Array(await blob.slice(offset+header+8,offset+header+12).arrayBuffer());handlers.add(String.fromCharCode(...value));}
          if(['moov','trak','mdia'].includes(type))await boxes(offset+header,offset+size,depth+1);
          offset+=size;
        }
      }
      await boxes(0,blob.size,0);return handlers.has('vide')&&mediaData;
    }
    async function fetchResource(url,job,kind){
      checkDownload(job);const safe=resourceURL(url);if(!safe)throw Error('未找到当前笔记的下载地址，请重新读取附件');
      const controller=new AbortController();let stallTimer,timedOut=false,onAbort,blob;
      const relayAbort=()=>controller.abort();job.controller.signal.addEventListener('abort',relayAbort,{once:true});
      const aborted=new Promise((_,reject)=>{onAbort=()=>reject(job.controller.signal.aborted?new DOMException('Download cancelled','AbortError'):Error(timedOut?'资源响应超时，请重试':'资源请求已停止'));controller.signal.addEventListener('abort',onAbort,{once:true});});
      const wait=promise=>{clearTimeout(stallTimer);stallTimer=setTimeout(()=>{timedOut=true;controller.abort();},20000);return Promise.race([promise,aborted]).finally(()=>clearTimeout(stallTimer));};
      try{
        const response=await wait(fetch(safe,{credentials:'omit',signal:controller.signal}));if(!response.ok)throw Error('资源请求失败（HTTP '+response.status+'），请重试');
        const label=downloadMessage,total=Number(response.headers.get('content-length'))||0;
        if(response.body){const reader=response.body.getReader(),chunks=[];let received=0,lastUpdate=0;
          try{while(true){checkDownload(job);const {done,value}=await wait(reader.read());if(done)break;chunks.push(value);received+=value.byteLength;
            if(Date.now()-lastUpdate>200){lastUpdate=Date.now();downloadProgress(job,label+' '+(received/1048576).toFixed(1)+' MB'+(total?' / '+(total/1048576).toFixed(1)+' MB':''));}
          }}finally{if(controller.signal.aborted)reader.cancel().catch(()=>{});reader.releaseLock();}
          blob=new Blob(chunks,{type:response.headers.get('content-type')||'application/octet-stream'});
        }else blob=await wait(response.blob());
      }finally{clearTimeout(stallTimer);job.controller.signal.removeEventListener('abort',relayAbort);controller.signal.removeEventListener('abort',onAbort);}
      checkDownload(job);const h=new Uint8Array(await blob.slice(0,32).arrayBuffer());
      let extension='';if(h[0]===255&&h[1]===216&&h[2]===255)extension='jpg';else if(h[0]===137&&h[1]===80&&h[2]===78&&h[3]===71&&h[4]===13&&h[5]===10&&h[6]===26&&h[7]===10)extension='png';
      else if(String.fromCharCode(...h.slice(0,4))==='RIFF'&&String.fromCharCode(...h.slice(8,12))==='WEBP')extension='webp';
      else if(/^GIF8[79]a$/.test(String.fromCharCode(...h.slice(0,6))))extension='gif';
      else if(String.fromCharCode(...h.slice(4,8))==='ftyp'){
        const brands=String.fromCharCode(...h.slice(8));
        if(/avif|avis/.test(brands))extension='avif';
        else if(kind==='video'&&await mp4HasTracks(blob,job))extension='mp4';
      }
      checkDownload(job);
      if(!blob.size||!extension||(kind==='video'&&extension!=='mp4'))throw Error('返回内容不是可下载的'+(kind==='video'?'MP4 文件':'图片')+'，已停止保存');
      if(kind==='image'){let bitmap;try{bitmap=await createImageBitmap(blob);if(!bitmap.width||!bitmap.height)throw Error('Empty image');}catch{throw Error('图片内容不完整或无法解码，已停止保存');}finally{bitmap?.close();}checkDownload(job);}
      return {blob,extension};
    }
    function saveResource(blob,name,job){
      checkDownload(job);if(downloadReady){URL.revokeObjectURL(downloadReady.url);downloadURLs.delete(downloadReady.url);}
      const url=URL.createObjectURL(blob);downloadURLs.add(url);downloadReady={url,name:safeFilename(name),key:job.key};
      const anchor=document.createElement('a');anchor.href=url;anchor.download=downloadReady.name;anchor.hidden=true;shadow.append(anchor);anchor.click();anchor.remove();
    }
    function zipImages(entries){
      if(entries.length>65535)throw Error('图片数量超过 ZIP 支持范围');
      const crcTable=Uint32Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
      const crc32=bytes=>{let crc=0xffffffff;for(const byte of bytes)crc=crcTable[(crc^byte)&255]^(crc>>>8);return (crc^0xffffffff)>>>0;};
      const parts=[],central=[];let offset=0,centralLength=0;const now=new Date(),time=(now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1),date=((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
      for(const entry of entries){const name=new TextEncoder().encode(entry.name),bytes=entry.bytes,crc=crc32(bytes);if(bytes.length>0xffffffff||offset>0xffffffff)throw Error('图片总大小超过 ZIP 支持范围');
        const local=new Uint8Array(30+name.length),l=new DataView(local.buffer);l.setUint32(0,0x04034b50,true);l.setUint16(4,20,true);l.setUint16(6,0x800,true);l.setUint16(10,time,true);l.setUint16(12,date,true);l.setUint32(14,crc,true);l.setUint32(18,bytes.length,true);l.setUint32(22,bytes.length,true);l.setUint16(26,name.length,true);local.set(name,30);parts.push(local,bytes);
        const header=new Uint8Array(46+name.length),c=new DataView(header.buffer);c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x800,true);c.setUint16(12,time,true);c.setUint16(14,date,true);c.setUint32(16,crc,true);c.setUint32(20,bytes.length,true);c.setUint32(24,bytes.length,true);c.setUint16(28,name.length,true);c.setUint32(42,offset,true);header.set(name,46);central.push(header);centralLength+=header.length;offset+=local.length+bytes.length;
      }
      if(offset>0xffffffff||centralLength>0xffffffff||offset+centralLength>0xffffffff)throw Error('图片总大小超过 ZIP 支持范围');
      const end=new Uint8Array(22),e=new DataView(end.buffer);e.setUint32(0,0x06054b50,true);e.setUint16(8,entries.length,true);e.setUint16(10,entries.length,true);e.setUint32(12,centralLength,true);e.setUint32(16,offset,true);return new Blob([...parts,...central,end],{type:'application/zip'});
    }
    async function downloadResource(kind){
      if(downloadJob||!detail)return;cancelDownload(true);const snapshot={key:detail.key,title:detail.title,images:[...detail.images],videos:[...detail.videoFiles],index:gallery};
      const job={key:snapshot.key,controller:new AbortController()};downloadJob=job;downloadKey=job.key;downloadMessage='正在准备下载…';updateDownloadUI();
      try{const filename=resourceBasename(snapshot);
        if(kind==='video'){
          if(!snapshot.videos.length)throw Error('当前笔记尚未提供可下载的 MP4 文件，请重新读取附件');let result,lastError;
          for(const url of [...new Set(snapshot.videos)]){checkDownload(job);try{downloadProgress(job,'正在下载视频…');result=await fetchResource(url,job,'video');break;}catch(error){if(error.name==='AbortError')throw error;lastError=error;}}
          if(!result)throw lastError||Error('视频下载失败');saveResource(result.blob,filename+'.mp4',job);
        }else if(kind==='image'){
          downloadProgress(job,'正在下载第 '+(snapshot.index+1)+' 张图片…');const result=await fetchResource(snapshot.images[snapshot.index],job,'image');saveResource(result.blob,filename+'_'+String(snapshot.index+1).padStart(2,'0')+'.'+result.extension,job);
        }else{
          if(!snapshot.images.length)throw Error('当前笔记没有已读取的图片');const entries=[];
          for(let i=0;i<snapshot.images.length;i++){downloadProgress(job,'正在打包图片 '+(i+1)+' / '+snapshot.images.length+'…');const result=await fetchResource(snapshot.images[i],job,'image');entries.push({name:filename+'_'+String(i+1).padStart(2,'0')+'.'+result.extension,bytes:new Uint8Array(await result.blob.arrayBuffer())});}
          downloadProgress(job,'正在生成 ZIP…');saveResource(zipImages(entries),filename+'_图片.zip',job);
        }
        downloadMessage='已提交浏览器保存；若未开始，请点击「保存文件」';
      }catch(error){if(downloadJob===job)downloadMessage=error.name==='AbortError'?'已取消下载':error.message==='Failed to fetch'?'资源下载失败，请重试或查看原文':error.message;}
      finally{if(downloadJob===job){downloadJob=null;updateDownloadUI();}}
    }
    function readDetail() {
      const el = document.querySelector("#noteContainer");
      if (!el) return null;
      const key =
        document.querySelector(".note-detail-mask")?.getAttribute("note-id") ||
        location.pathname.match(/\/(?:explore|discovery\/item)\/([^/]+)/)?.[1] || location.pathname;
      if (closing && key === closeKey) return null;
      const slides = [
        ...el.querySelectorAll(
          ".media-container .swiper-slide:not(.swiper-slide-duplicate)",
        ),
      ];
      let images = (slides.length ? slides : [el.querySelector(".media-container")].filter(Boolean)).map(s=>resourceURL(s?.querySelector('img')?.src));
      const note=pageNote(key);
      if(note?.imageList?.length)images=note.imageList.map((image,index)=>[image?.urlDefault,...(image?.infoList||[]).filter(info=>info?.imageScene==='WB_DFT').map(info=>info.url),...(image?.infoList||[]).map(info=>info?.url),images[index]].map(resourceURL).find(Boolean)||'');
      const nativeVideo=note?.type==='normal'?null:videoLease?.key===key?videoLease.video:el.querySelector('video');
      return {
        el,
        key,
        videoFiles:noteVideoFiles(note,nativeVideo),
        originalUrl:originalNoteURL(key,note),
        title: text(el, "#detail-title") || note?.title || cache.get(key)?.title || (text(el,"#detail-desc") || note?.desc || '').split('\n')[0].slice(0,80) || (note?'无标题笔记':'加载中…'),
        author: text(el, ".username"),
        body: text(el, "#detail-desc"),
        date: text(el, ".note-content .date") || text(el, ".date"),
        images,
        video:nativeVideo,
        total: text(el, ".comments-container .total") || text(el, ".total"),
        end: !!el.querySelector(".end-container,.no-more,.comments-end"),
        comments: [...el.querySelectorAll(".comment-item")]
          .map((c, i) => {
            const sub = c.classList.contains("comment-item-sub");
            const more = !sub
              ? c.parentElement.querySelector(".reply-container > .show-more")
              : null;
            return {
              id: c.id || String(i),
              el: c,
              sub,
              author: text(c, ".author .name") || text(c, ".name"),
              body:
                text(c, ".content") ||
                (c.querySelector(".content img") ? "（图片或表情评论）" : ""),
              date: text(c, ".date"),
              likes: text(c, ".like-wrapper .count"),
              more,
              moreText: more?.textContent.trim() || "",
            };
          })
          .filter((c) => c.body),
      };
    }
    function setDetail(on) {
      $("#list").hidden = on;
      $("#detail").hidden = !on;
      $("#detail-tab").hidden = !on;
      $(".app").classList.toggle("detail-view", on);
      $('.sheet[data-action="list"]').classList.toggle("current", !on);
      $("#detail-tab").classList.toggle("current", on);
    }
    function detailNav() {
      const a = filtered(),
        i = a.findIndex((r) => r.id === activeId);
      $("#note-prev").disabled = i <= 0 || pending;
      $("#note-next").disabled = i < 0 || i >= a.length - 1 || pending;
      $("#note-position").textContent = i >= 0 ? i + 1 + " / " + a.length : "";
    }
    function renderDetail(d) {
      releaseVideo();
      $("#detail").innerHTML =
        '<div class="detail-head"><button data-action="back">← 返回台账</button><button data-action="note-prev" id="note-prev">上一条</button><span class="position" id="note-position"></span><button data-action="note-next" id="note-next">下一条</button><span class="grow"></span><a class="original-link" data-action="original" href="'+esc(d.originalUrl)+'" target="_blank" rel="noopener noreferrer">查看原文 ↗</a><button data-action="media-jump">附件</button><button data-action="comments-jump">评论</button><button data-action="top">回到正文</button></div><table class="detail-table"><colgroup><col style="width:42px"><col style="width:112px"><col></colgroup><tbody><tr><td class="num">1</td><td class="label">标题</td><td><strong>' +
        esc(d.title || "加载中…") +
        '</strong></td></tr><tr><td class="num">2</td><td class="label">作者 / 日期</td><td>' +
        esc(d.author) +
        '　<span style="color:#888">' +
        esc(d.date) +
        '</span></td></tr><tr id="attachment-row"><td class="num">3</td><td class="label">附件</td><td><button data-action="media" id="media-toggle"></button><div id="media-slot"></div></td></tr><tr id="body-row"><td class="num">4</td><td class="label">正文</td><td><div class="prose readable">' +
        esc(d.body || "（暂无正文）") +
        '</div></td></tr></tbody></table><div class="comments-heading" id="comments-heading"><strong>评论记录</strong><span id="comment-count"></span></div><table class="comment-table"><colgroup><col style="width:42px"><col style="width:190px"><col><col style="width:140px"></colgroup><thead><tr><th></th><th>用户</th><th>评论内容</th><th>日期 / 点赞</th></tr></thead><tbody id="comment-rows"></tbody></table><div class="comment-more"><button data-action="comments">继续加载评论</button><span class="load-state" role="status"></span></div>';
      detailNav();
      renderMedia();
      renderComments(d);
    }
    function renderMedia() {
      if (!detail || !$("#media-slot")) return;
      const toggle = $("#media-toggle");
      toggle.textContent = mediaOpen
        ? "收起附件"
        : detail.video
          ? "展开视频"
          : "展开图片（" + detail.images.length + "）";
      const slot = $("#media-slot");
      if (
        mediaOpen &&
        !detail.video &&
        detail.images.length &&
        slot.querySelector(".media-stage img")
      ) {
        gallery = Math.max(0, Math.min(gallery, detail.images.length - 1));
        const img = slot.querySelector("img");
        img.src = detail.images[gallery];
        img.alt = "笔记图片 " + (gallery + 1);
        slot.querySelector(".media-tools span").textContent =
          gallery + 1 + " / " + detail.images.length;
        slot.querySelector('[data-action="prev"]').disabled = gallery === 0;
        slot.querySelector('[data-action="next"]').disabled =
          gallery === detail.images.length - 1;
        downloadToolbar();
        return;
      }
      releaseVideo();
      slot.replaceChildren();
      zoomed = false;
      if (!mediaOpen) return;
      const box = document.createElement("div");
      box.className = "media";
      slot.append(box);
      if (detail.video) {
        box.innerHTML =
          '<div class="video-tools"><button data-action="video-toggle">播放</button><input class="video-progress" id="video-progress" type="range" aria-label="播放进度" min="0" max="0" step="0.1" value="0"><span class="video-time" id="video-time"></span><button data-action="video-mute">静音</button><input class="video-volume" id="video-volume" type="range" aria-label="视频音量" min="0" max="100" step="1" value="' +
          Math.round(soundVolume * 100) +
          '"><select id="video-rate" aria-label="播放速度"><option value="0.75">0.75 倍</option><option value="1" selected>1 倍</option><option value="1.25">1.25 倍</option><option value="1.5">1.5 倍</option><option value="2">2 倍</option></select></div><div class="media-stage video-stage"></div><div class="video-help"><span id="video-status">已暂停</span> · 空格播放 / 暂停 · ← → 快进 / 后退 5 秒 · 遮罩或切换笔记后需手动继续</div>';
        mountVideo(box.querySelector(".video-stage"), detail.video);
        downloadToolbar();
        return;
      }
      if (!detail.images.length) {
        box.innerHTML =
          '<div style="padding:20px">图片尚未加载。<button data-action="media-retry">重新读取</button></div>';
        return;
      }
      gallery = Math.max(0, Math.min(gallery, detail.images.length - 1));
      box.innerHTML =
        '<div class="media-tools"><button data-action="prev" ' +
        (gallery === 0 ? "disabled" : "") +
        ">上一张</button><span>" +
        (gallery + 1) +
        " / " +
        detail.images.length +
        '</span><button data-action="next" ' +
        (gallery === detail.images.length - 1 ? "disabled" : "") +
        '>下一张</button><button data-action="zoom">放大阅读</button><span style="color:#777;font-size:11px">← → 翻页</span></div><div class="media-stage"><img src="' +
        esc(detail.images[gallery]) +
        '" alt="笔记图片 ' +
        (gallery + 1) +
        '"></div>';
      downloadToolbar();
      box
        .querySelector("img")
        .addEventListener(
          "error",
          () => status("图片加载失败，可重新读取附件"),
          { once: true },
        );
    }
    function commentHTML(c, i) {
      return (
        '<tr data-comment-id="' +
        esc(c.id) +
        '" class="' +
        (c.sub ? "reply-row" : "") +
        '"><td class="num">' +
        (i + 1) +
        "</td><td>" +
        (c.sub ? "↳　" : "") +
        esc(c.author) +
        '</td><td class="comment-body">' +
        esc(c.body) +
        (c.more
          ? '<button class="reply-toggle" data-action="replies" data-comment="' +
            esc(c.id) +
            '">' +
            esc(c.moreText) +
            "</button>"
          : "") +
        '</td><td style="color:#777">' +
        esc(c.date) +
        "<br>" +
        esc(c.likes) +
        "</td></tr>"
      );
    }
    function renderComments(d) {
      const tbody = $("#comment-rows");
      if (!tbody) return;
      const top = viewport.scrollTop;
      const old = [...tbody.children];
      const pool = new Map(old.map((e) => [e.dataset.commentId, e]));
      d.comments.forEach((c, i) => {
        const html = commentHTML(c, i);
        let row = pool.get(c.id);
        if (!row) {
          const temp = document.createElement("tbody");
          temp.innerHTML = html;
          row = temp.firstElementChild;
        } else {
          const temp = document.createElement("tbody");
          temp.innerHTML = html;
          const fresh = temp.firstElementChild;
          if (row.innerHTML !== fresh.innerHTML) {
            row.innerHTML = fresh.innerHTML;
            row.className = fresh.className;
          }
        }
        if (tbody.children[i] !== row)
          tbody.insertBefore(row, tbody.children[i] || null);
        pool.delete(c.id);
      });
      pool.forEach((row) => row.remove());
      viewport.scrollTop = top;
      $("#comment-count").textContent =
        (d.total ? d.total + " · " : "") +
        "已显示 " +
        d.comments.length +
        " 条（含回复）";
      updateCommentState();
    }
    function updateCommentState() {
      const b = $(".comment-more button"),
        s = $(".comment-more .load-state");
      if (!b) return;
      b.disabled = commentsBusy || commentsEnd;
      b.textContent = commentsBusy
        ? "正在加载评论…"
        : commentsEnd
          ? "评论列表已到底"
          : isLoggedOut()
            ? "登录后加载更多评论"
          : commentsFailed
            ? "重试加载评论"
            : "继续加载评论";
      s.textContent = commentsBusy
        ? "正在读取下一批，当前阅读位置保持不变"
        : commentsEnd
          ? "主评论已加载完，未展开的回复可在对应评论下查看"
          : commentsFailed
            ? isLoggedOut() ? "当前未登录，请登录后继续读取评论。" : "暂未收到新评论，可重试或切回原界面检查。"
            : "在评论末尾继续阅读；向下滚动可加载下一批";
    }
    function loadComments() {
      if (!detail || commentsBusy || commentsEnd || panic || !enabled) return;
      if (isLoggedOut()) { requestLogin('加载更多评论'); return; }
      commentsBusy = true;
      commentsFailed = false;
      commentsBefore = detail.comments.length;
      const generation = ++commentsGeneration,
        key = detail.key;
      updateCommentState();
      const scroller = detail.el.querySelector(".note-scroller");
      if (!scroller) {
        commentsBusy = false;
        commentsFailed = true;
        updateCommentState();
        return;
      }
      scroller.scrollTop = scroller.scrollHeight;
      scroller.dispatchEvent(new Event("scroll"));
      commentsTimer = later(() => {
        if (detail?.key !== key || generation !== commentsGeneration) return;
        commentsBusy = false;
        commentsFailed = true;
        updateCommentState();
      }, 7000);
    }
    function back() {
      ++openSerial;
      pending = false;
      wantMedia = false;
      mediaOpen = false;
      commentsGeneration++;
      closing = true;
      closeKey = detail?.key || "";
      closeNative();
      detail = null;
      detailKey = "";
      lastBodySignature = "";
      lastCommentsSignature = "";
      setDetail(false);
      viewport.scrollTop = listTop;
      select(filtered().findIndex((r) => r.id === activeId));
      status("就绪");
      later(() => {
        closing = false;
      }, 800);
    }
    function more() {
      if (detail || feedBusy || feedEnd || panic) return;
      if (feedFailed && isLoggedOut()) { requestLogin('继续加载记录'); return; }
      const feed =
        document.querySelector("#exploreFeeds") ||
        document.querySelector("section.note-item")?.parentElement;
      if (!feed) {
        status("当前页面没有可加载的列表");
        return;
      }
      feedBusy = true;
      feedFailed = false;
      feedBefore = cache.size;
      updateFeedState();
      const last = feed.querySelector("section.note-item:last-of-type");
      last?.scrollIntoView({ block: "end" });
      window.scrollTo(0, document.documentElement.scrollHeight);
      let p = feed.parentElement;
      while (p && p !== document.body) {
        if (
          p.scrollHeight > p.clientHeight + 100 &&
          /auto|scroll/.test(getComputedStyle(p).overflowY)
        )
          p.scrollTop = p.scrollHeight;
        p = p.parentElement;
      }
      feedTimer = later(() => {
        if (feedBusy) {
          feedBusy = false;
          feedFailed = true;
          updateFeedState();
        }
      }, 7000);
    }
    function jump(selector) {
      const el = $(selector);
      if (el)
        viewport.scrollTop +=
          el.getBoundingClientRect().top -
          viewport.getBoundingClientRect().top -
          50;
    }
    function togglePanic() {
      if(!panic)cancelDownload(true);
      panic = !panic;
      document.body.classList.toggle('xw-panic', panic);
      pause();
      $(".app").classList.toggle("mask-view", panic);
      $("#mask").hidden = !panic;
      if (panic) {
        maskScroll = viewport.scrollTop;
        savedSearch = $(".search").value;
        savedFilter = $("#filter").value;
        $(".search").value = "";
        $("#filter").value = "";
        $("#list").hidden = true;
        $("#detail").hidden = true;
        document.title = "月度工作汇总.xlsx";
        $(".filename").textContent = "月度工作汇总.xlsx";
        $(".formula-value").textContent = "月度项目进度统计";
        $("#total").textContent = "24 条记录";
        $("#mask").innerHTML =
          "<table>" +
          col +
          '<thead><tr><th></th><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th></tr></thead><tbody><tr class="headers"><td class="num">1</td><td>序号</td><td>部门</td><td>项目名称</td><td>负责人</td><td>完成率</td><td>状态</td></tr>' +
          Array.from(
            { length: 24 },
            (_, i) =>
              '<tr><td class="num">' +
              (i + 2) +
              "</td><td>" +
              String(i + 1).padStart(3, "0") +
              "</td><td>业务支持</td><td>" +
              ["月度数据归档", "需求清单整理", "项目资料核对", "工作进度汇总"][
                i % 4
              ] +
              "</td><td>项目组</td><td>" +
              [100, 85, 70, 100][i % 4] +
              "%</td><td>" +
              (i % 3 ? "进行中" : "已完成") +
              "</td></tr>",
          ).join("") +
          "</tbody></table>";
        viewport.scrollTop = 0;
        status("就绪");
      } else {
        $(".search").value = savedSearch;
        $("#filter").value = savedFilter;
        $(".filename").textContent = "素材整理.xlsx";
        setDetail(!!detail);
        sync();
        $("#total").textContent = filtered().length + " 条记录";
        $(".formula-value").textContent =
          detail?.title ||
          filtered()[selected]?.title ||
          "点击标题阅读笔记，点击附件直接看图";
        viewport.scrollTop = maskScroll;
      }
    }
    function sync() {
      if (enabled) { updateWorkbookIcon(); syncAuthentication(); }
      if (!enabled || panic) return;
      const d = readDetail();
      const incoming = readRows();
      const nextContext = location.pathname.includes("search_result")
        ? location.pathname + location.search
        : "/explore";
      if (!d && !pending && !closing && nextContext !== context) {
        context = nextContext;
        cache.clear();
        activeId = "";
        selected = -1;
        filter = "";
        $("#filter").value = "";
        feedBusy = false;
        feedFailed = false;
        feedEnd = false;
        signature = "";
        listTop = 0;
        viewport.scrollTop = 0;
        $(".formula-value").textContent = "点击标题阅读笔记，点击附件直接看图";
      }
      incoming.forEach((r) => cache.set(r.id, r));
      rows = [...cache.values()];
      const sig =
        JSON.stringify(
          rows.map((r) => [r.id, r.title, r.likes, r.author, r.cover]),
        ) +
        thumb +
        filter +
        sort;
      if (sig !== signature) {
        signature = sig;
        renderList();
      }
      if (feedBusy && cache.size > feedBefore) {
        feedBusy = false;
        feedFailed = false;
        clearTimeout(feedTimer);
        updateFeedState();
      }
      if (d) {
        const newNote = d.key !== detailKey;
        if(newNote)cancelDownload(true);
        detail = d;
        if (newNote) {
          detailKey = d.key;
          pending = false;
          closing = false;
          if (!activeId && cache.has(d.key)) {
            activeId = d.key;
            selected = filtered().findIndex((r) => r.id === activeId);
          }
          mediaOpen = wantMedia;
          wantMedia = false;
          gallery = 0;
          zoomed = false;
          commentsBusy = false;
          commentsFailed = false;
          commentsEnd = false;
          commentsGeneration++;
          lastLoadScroll = -1;
          lastBodySignature = "";
          lastCommentsSignature = "";
          viewport.scrollTop = 0;
        }
        const ds = JSON.stringify([
          d.key,
          d.title,
          d.body,
          d.author,
          d.date,
          !!d.video,
          d.originalUrl,
          d.images,
          d.videoFiles,
        ]);
        if (ds !== lastBodySignature) {
          lastBodySignature = ds;
          renderDetail(d);
          if (mediaOpen) jump("#attachment-row");
        }
        if (mediaOpen && d.video && !videoLease && $("#media-slot"))
          renderMedia();
        const cs = JSON.stringify(
          d.comments.map((c) => [c.id, c.author, c.body, c.likes, c.moreText]),
        );
        if (cs !== lastCommentsSignature) {
          lastCommentsSignature = cs;
          renderComments(d);
        }
        if (commentsBusy && d.comments.length > commentsBefore) {
          commentsBusy = false;
          commentsFailed = false;
          clearTimeout(commentsTimer);
          updateCommentState();
        }
        if (d.end) {
          commentsEnd = true;
          commentsBusy = false;
          updateCommentState();
        }
        setDetail(true);
        detailNav();
        $(".formula-value").textContent = d.title;
        $(".address").textContent = "C1";
        status("记录详情");
        mediaElements().forEach((v) => {
          if (v !== videoLease?.video || !canPlay(v)) {
            v.muted = true;
            if (!v.paused) v.pause();
          }
        });
      } else if (detail && !pending) {
        cancelDownload(true);
        detail = null;
        detailKey = "";
        lastBodySignature = "";
        setDetail(false);
        viewport.scrollTop = listTop;
        status("就绪");
      }
      document.title = "素材整理.xlsx";
      if (loginDialog()) status(authMessage);
      currentRoute = location.pathname + location.search;
    }
    shadow.addEventListener("click", (e) => {
      const b = e.target.closest("[data-action]");
      if (b) {
        const a = b.dataset.action;
        if (panic && a !== "panic" && a !== "native") return;
        const actions = {
          home: () => {
            if (detail) back();
            document.querySelector("#link-guide")?.click();
          },
          sync: () => {
            signature = "";
            sync();
            status("记录已同步");
          },
          more,
          open: () => open(+b.dataset.index),
          attachment: () => open(+b.dataset.index, true),
          thumb: () => {
            thumb = !thumb;
            b.classList.toggle("pressed", thumb);
            signature = "";
            sync();
          },
          sort: () => {
            sort = !sort;
            b.classList.toggle("pressed", sort);
            signature = "";
            sync();
          },
          native,
          login: () => requestLogin('账户登录'),
          panic: togglePanic,
          back,
          list: () => {
            if (detail) back();
          },
          detail: () => {},
          media: () => {
            mediaOpen = !mediaOpen;
            if (!mediaOpen) pause();
            renderMedia();
          },
          "media-retry": () => {
            detail = readDetail() || detail;
            renderMedia();
          },
          "media-jump": () => {
            if (!mediaOpen) {
              mediaOpen = true;
              renderMedia();
            }
            jump("#attachment-row");
          },
          "comments-jump": () => jump("#comments-heading"),
          top: () => {
            viewport.scrollTop = 0;
          },
          prev: () => {
            gallery--;
            renderMedia();
          },
          next: () => {
            gallery++;
            renderMedia();
          },
          zoom: () => {
            zoomed = !zoomed;
            const stage = $(".media-stage");
            stage.classList.toggle("zoom", zoomed);
            stage.scrollTop = 0;
            stage.scrollLeft = 0;
            b.textContent = zoomed ? "适应窗口" : "放大阅读";
          },
          original:()=>pause(),
          "download-image":()=>downloadResource("image"),
          "download-all":()=>downloadResource("all"),
          "download-video":()=>downloadResource("video"),
          "download-cancel":()=>cancelDownload(),
          "video-toggle": toggleVideo,
          "video-mute": toggleSound,
          comments: loadComments,
          replies: () => {
            if (isLoggedOut()) { requestLogin('展开更多回复'); return; }
            const c = detail?.comments.find((c) => c.id === b.dataset.comment);
            if (c?.more) {
              b.disabled = true;
              b.textContent = "正在展开…";
              c.more.click();
              later(() => {
                if (b.isConnected) {
                  b.disabled = false;
                  b.textContent = c.moreText;
                }
              }, 7000);
            }
          },
          "note-prev": () => open(selected - 1),
          "note-next": () => open(selected + 1),
        };
        actions[a]?.();
        return;
      }
      const row = e.target.closest("[data-row]");
      if (row) select(+row.dataset.row);
    });
    shadow.addEventListener('auxclick',e=>{if(e.button===1&&e.target.closest('[data-action="original"]'))pause();});
    shadow.addEventListener("dblclick", (e) => {
      if (e.target.closest("button")) return;
      const row = e.target.closest("[data-row]");
      if (row && !panic) open(+row.dataset.row);
    });

    shadow.addEventListener("input", (event) => {
      const v = videoLease?.video;
      if (!v) return;
      if (event.target.id === "video-progress") {
        const time = Number(event.target.value);
        if (Number.isFinite(v.duration))
          v.currentTime = Math.max(0, Math.min(time, v.duration));
        updateVideoUI();
      }
      if (event.target.id === "video-volume") {
        soundVolume = Number(event.target.value) / 100;
        if (soundVolume > 0) soundMuted = false;
        v.volume = soundVolume;
        v.muted = soundMuted || soundVolume === 0;
        updateVideoUI();
      }
    });
    shadow.addEventListener("change", (event) => {
      if (event.target.id === "video-rate" && videoLease)
        videoLease.video.playbackRate = Number(event.target.value);
    });
    $("#filter").addEventListener("input", (e) => {
      filter = e.target.value;
      signature = "";
      sync();
      viewport.scrollTop = 0;
    });
    $("#search-form").addEventListener("submit", (e) => {
      e.preventDefault();
      if (panic) return;
      const keyword = $(".search").value.trim();
      if (!keyword) return;
      if (isLoggedOut()) { requestLogin('搜索'); return; }
      pause();
      const input =
        document.querySelector("#search-input") ||
        document.querySelector("#search-input-in-feeds,#search-input-ai");
      if (!input) {
        status("未找到站内搜索框，可切回原界面搜索");
        return;
      }
      const proto =
        input.tagName === "TEXTAREA"
          ? HTMLTextAreaElement.prototype
          : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, "value").set.call(input, keyword);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      const submit =
        input.parentElement.querySelector(".search-icon") ||
        input
          .closest(".textarea-container")
          ?.querySelector(".submit-button-image");
      if (submit) later(() => submit.click(), 60);
      else
        input.dispatchEvent(
          new KeyboardEvent("keydown", {
            key: "Enter",
            code: "Enter",
            keyCode: 13,
            bubbles: true,
          }),
        );
      status("正在搜索…");
    });
    function keys(e) {
      if (e.altKey && e.code === "KeyQ") {
        e.preventDefault();
        e.stopPropagation();
        if (!enabled) show();
        togglePanic();
        return;
      }
      if (e.altKey && e.code === "KeyW") {
        e.preventDefault();
        e.stopPropagation();
        native();
        return;
      }
      if (!enabled || panic || loginDialog()) return;
      if (
        e
          .composedPath()
          .some(
            (n) =>
              n instanceof HTMLInputElement ||
              n instanceof HTMLTextAreaElement ||
              n instanceof HTMLSelectElement,
          )
      )
        return;
      if (
        detail &&
        mediaOpen &&
        videoLease &&
        ["Space", "ArrowLeft", "ArrowRight"].includes(e.code)
      ) {
        if (
          e.composedPath().some((n) => n instanceof HTMLButtonElement) &&
          e.code === "Space"
        )
          return;
        e.preventDefault();
        if (e.code === "Space") toggleVideo();
        else
          videoLease.video.currentTime = Math.max(
            0,
            Math.min(
              videoLease.video.duration || 0,
              videoLease.video.currentTime + (e.code === "ArrowRight" ? 5 : -5),
            ),
          );
        return;
      }
      if (e.code === "Escape" && detail) {
        e.preventDefault();
        back();
        return;
      }
      if (
        detail &&
        mediaOpen &&
        !detail.video &&
        ["ArrowLeft", "ArrowRight"].includes(e.key)
      ) {
        e.preventDefault();
        gallery += e.key === "ArrowRight" ? 1 : -1;
        renderMedia();
        return;
      }
      if (!detail && ["ArrowDown", "ArrowUp", "Enter"].includes(e.key)) {
        e.preventDefault();
        if (e.key === "Enter") open(selected);
        else {
          select(
            Math.max(
              0,
              Math.min(
                filtered().length - 1,
                selected + (e.key === "ArrowDown" ? 1 : -1),
              ),
            ),
          );
          shadow
            .querySelector('tr[data-row="' + selected + '"]')
            ?.scrollIntoView({ block: "nearest" });
        }
      }
    }
    document.addEventListener("keydown", keys, true);
    let prevScroll = 0;
    viewport.addEventListener(
      "scroll",
      () => {
        const y = viewport.scrollTop,
          down = y > prevScroll + 2;
        prevScroll = y;
        if (!down || !enabled || panic || pending) return;
        if (detail) {
          if (isLoggedOut()) return;
          if (commentsBusy || commentsFailed || commentsEnd) return;
          const end = $(".comment-more");
          if (
            end &&
            end.getBoundingClientRect().top <
              viewport.getBoundingClientRect().bottom + 80 &&
            y > lastLoadScroll + 80
          ) {
            lastLoadScroll = y;
            loadComments();
          }
        } else if (
          !feedBusy &&
          !feedFailed &&
          !feedEnd &&
          !filter &&
          viewport.scrollHeight - y - viewport.clientHeight < 120
        )
          more();
      },
      { passive: true },
    );
    const observer = new MutationObserver(() => {
      if (scheduleTimer) return;
      scheduleTimer = later(() => {
        scheduleTimer = 0;
        sync();
      }, 150);
    });
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true,
    });
    const timer = setInterval(sync, 1800);
    window.__xhsWorkbook = {
      version: "0.4.5",
      show,
      destroy() {
        iconObserver.disconnect();
        restoreOriginalIcons();
        cancelDownload(true);
        downloadURLs.forEach(url=>URL.revokeObjectURL(url));
        downloadURLs.clear();
        pause();
        releaseVideo(true);
        document.removeEventListener("play", guardMedia, true);
        document.removeEventListener("playing", guardMedia, true);
        shadow.removeEventListener("play", guardMedia, true);
        shadow.removeEventListener("playing", guardMedia, true);
        document.removeEventListener("visibilitychange", pauseHidden);
        ++openSerial;
        commentsGeneration++;
        observer.disconnect();
        clearInterval(timer);
        timers.forEach(clearTimeout);
        document.removeEventListener("keydown", keys, true);
        document.body.classList.remove("xw-active", "xw-panic");
        document.title = originalTitle;
        host.remove();
        globalStyle.remove();
        delete window.__xhsWorkbook;
      },
    };
    show();
  }
})();
