<%@ Page Language="C#" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<WebPartPages:AllowFraming runat="server" />
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<meta name="robots" content="noindex, nofollow" />
<title>SharePoint Calendar</title>
<style>
/* ===== CSS Custom Properties ===== */
:root {
  --color-primary: #0078d4;
  --color-primary-hover: #106ebe;
  --color-primary-light: #deecf9;
  --color-danger: #d13438;
  --color-danger-hover: #a4262c;
  --color-success: #107c10;
  --color-warning: #ffb900;
  --color-bg: #ffffff;
  --color-bg-secondary: #f3f2f1;
  --color-bg-tertiary: #edebe9;
  --color-surface: #ffffff;
  --color-surface-hover: #f3f2f1;
  --color-border: #e1dfdd;
  --color-border-strong: #c8c6c4;
  --color-text: #323130;
  --color-text-secondary: #605e5c;
  --color-text-tertiary: #a19f9d;
  --color-text-inverse: #ffffff;
  --color-overlay: rgba(0,0,0,0.4);
  --color-card-shadow: rgba(0,0,0,0.08);
  --color-label-bg: #e1dfdd;
  --color-label-text: #323130;
  --font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
  --radius-sm: 4px;
  --radius-md: 8px;
  --radius-lg: 12px;
  --shadow-sm: 0 1px 2px var(--color-card-shadow);
  --shadow-md: 0 2px 8px var(--color-card-shadow);
  --shadow-lg: 0 4px 16px var(--color-card-shadow);
  --transition: 0.2s ease;
}

@media (prefers-color-scheme: dark) {
  :root {
    --color-primary: #2b88d8;
    --color-primary-hover: #3c9eef;
    --color-primary-light: #1b3a5c;
    --color-danger: #ef6950;
    --color-danger-hover: #ff8c75;
    --color-success: #54b054;
    --color-warning: #ffc83d;
    --color-bg: #1b1a19;
    --color-bg-secondary: #252423;
    --color-bg-tertiary: #323130;
    --color-surface: #252423;
    --color-surface-hover: #323130;
    --color-border: #3b3a39;
    --color-border-strong: #484644;
    --color-text: #f3f2f1;
    --color-text-secondary: #c8c6c4;
    --color-text-tertiary: #8a8886;
    --color-text-inverse: #1b1a19;
    --color-overlay: rgba(0,0,0,0.6);
    --color-card-shadow: rgba(0,0,0,0.3);
    --color-label-bg: #3b3a39;
    --color-label-text: #f3f2f1;
  }
}

/* ===== Reset & Base ===== */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html, body {
  height: 100%;
  font-family: var(--font-family);
  font-size: 14px;
  line-height: 1.5;
  color: var(--color-text);
  background: var(--color-bg);
}

body { display: flex; flex-direction: column; overflow: hidden; }

/* ===== Page Header ===== */
.page-header {
  display: flex;
  align-items: center;
  padding: 12px 20px;
  background: var(--color-surface);
  border-bottom: 1px solid var(--color-border);
  flex-shrink: 0;
  gap: 12px;
  min-height: 52px;
}
.page-header.embed-mode { display: none; }
.page-title {
  font-size: 20px;
  font-weight: 600;
  color: var(--color-text);
  flex: 1;
}
.header-actions { display: flex; align-items: center; gap: 8px; }

/* ===== Save Indicator ===== */
.save-indicator {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--color-text-secondary);
  opacity: 0;
  transition: opacity 0.3s ease;
  white-space: nowrap;
}
.save-indicator.visible { opacity: 1; }
.save-indicator .spinner {
  width: 14px; height: 14px;
  border: 2px solid var(--color-border);
  border-top-color: var(--color-primary);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}
.save-indicator .checkmark {
  color: var(--color-success);
  font-size: 16px;
  font-weight: 700;
}
@keyframes spin { to { transform: rotate(360deg); } }

/* ===== Buttons ===== */
.btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 16px;
  border: 1px solid transparent;
  border-radius: var(--radius-sm);
  font-family: var(--font-family);
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: background var(--transition), border-color var(--transition), color var(--transition);
  white-space: nowrap;
}
.btn:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
.btn-primary {
  background: var(--color-primary);
  color: var(--color-text-inverse);
}
.btn-primary:hover:not(:disabled) { background: var(--color-primary-hover); }
.btn-secondary {
  background: transparent;
  color: var(--color-primary);
  border-color: var(--color-primary);
}
.btn-secondary:hover:not(:disabled) { background: var(--color-primary-light); }
.btn-danger {
  background: transparent;
  color: var(--color-danger);
  border-color: var(--color-danger);
}
.btn-danger:hover:not(:disabled) {
  background: var(--color-danger);
  color: var(--color-text-inverse);
}
.btn-ghost {
  background: transparent;
  color: var(--color-text-secondary);
  border: none;
  padding: 6px 10px;
}
.btn-ghost:hover:not(:disabled) {
  background: var(--color-surface-hover);
  color: var(--color-text);
}
.btn-sm { font-size: 12px; padding: 4px 10px; }
.btn-icon {
  padding: 6px;
  border: none;
  background: transparent;
  color: var(--color-text-secondary);
  cursor: pointer;
  border-radius: var(--radius-sm);
  font-size: 18px;
  line-height: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}
.btn-icon:hover { background: var(--color-surface-hover); color: var(--color-text); }

/* ===== Main Layout ===== */
.main-container {
  display: flex;
  flex: 1;
  overflow: hidden;
}

/* ===== Sidebar (Card Panel) ===== */
.sidebar {
  width: 340px;
  min-width: 280px;
  max-width: 400px;
  border-right: 1px solid var(--color-border);
  display: flex;
  flex-direction: column;
  background: var(--color-bg-secondary);
  flex-shrink: 0;
}
.sidebar-header {
  padding: 12px 16px;
  border-bottom: 1px solid var(--color-border);
  font-size: 14px;
  font-weight: 600;
  color: var(--color-text-secondary);
  background: var(--color-surface);
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.sidebar-tabs {
  display: flex;
  background: var(--color-surface);
  border-bottom: 1px solid var(--color-border);
  padding: 0 4px;
  overflow-x: auto;
  scrollbar-width: none; /* Hide scrollbar for Firefox */
  -ms-overflow-style: none; /* Hide scrollbar for IE/Edge */
}
.sidebar-tabs::-webkit-scrollbar {
  display: none; /* Hide scrollbar for Chrome, Safari, Opera */
}
.sidebar-tab {
  flex: 0 0 auto;
  border: none;
  background: transparent;
  padding: 10px 12px;
  font-size: 11px;
  font-weight: 600;
  color: var(--color-text-tertiary);
  cursor: pointer;
  border-bottom: 2px solid transparent;
  text-align: center;
  transition: all var(--transition);
  white-space: nowrap;
}
.sidebar-tab:hover {
  color: var(--color-text-secondary);
}
.sidebar-tab.active {
  color: var(--color-primary);
  border-bottom-color: var(--color-primary);
}
.sidebar-scroll {
  flex: 1;
  overflow-y: auto;
  padding: 8px;
}
.sidebar-scroll::-webkit-scrollbar { width: 6px; }
.sidebar-scroll::-webkit-scrollbar-track { background: transparent; }
.sidebar-scroll::-webkit-scrollbar-thumb {
  background: var(--color-border-strong);
  border-radius: 3px;
}

/* ===== Cards ===== */
.card {
  background: var(--color-surface);
  border: 1px solid var(--color-border);
  border-radius: var(--radius-md);
  padding: 12px;
  margin-bottom: 8px;
  position: relative;
  transition: border-color var(--transition), box-shadow var(--transition);
  border-left: 3px solid var(--color-primary);
}
.card:hover { box-shadow: var(--shadow-md); }
.card.highlighted {
  border-color: var(--color-primary);
  box-shadow: 0 0 0 2px var(--color-primary-light), var(--shadow-md);
}
.card-field-heading {
  font-size: 15px;
  font-weight: 600;
  color: var(--color-text);
  margin-bottom: 2px;
}
.card-field-subheading {
  font-size: 13px;
  font-weight: 500;
  color: var(--color-text-secondary);
  margin-bottom: 4px;
}
.card-field-label {
  display: inline-block;
  background: var(--color-label-bg);
  color: var(--color-label-text);
  font-size: 11px;
  font-weight: 600;
  padding: 2px 8px;
  border-radius: 10px;
  margin: 2px 4px 2px 0;
}
.card-field-regular {
  font-size: 13px;
  color: var(--color-text-secondary);
  line-height: 1.4;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.card.expanded .card-field-regular {
  -webkit-line-clamp: unset;
  line-clamp: unset;
  display: block;
}
.card-dates {
  display: flex;
  gap: 8px;
  margin-top: 8px;
  font-size: 12px;
  color: var(--color-text-tertiary);
  align-items: center;
  flex-wrap: wrap;
}
.card-dates .date-label { font-weight: 600; }
.card-action {
  position: absolute;
  top: 8px;
  right: 8px;
  display: flex;
  gap: 4px;
}
.card-action .btn { font-size: 11px; padding: 3px 8px; }

/* Card date pickers (inline) */
.card-date-edit {
  margin-top: 8px;
  padding-top: 8px;
  border-top: 1px solid var(--color-border);
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.card-date-edit label {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-text-secondary);
}
.card-date-edit input[type="date"] {
  font-family: var(--font-family);
  font-size: 13px;
  padding: 4px 8px;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-sm);
  background: var(--color-bg);
  color: var(--color-text);
}
.card-date-edit .date-edit-actions {
  display: flex;
  gap: 6px;
  margin-top: 4px;
}

/* ===== Calendar Area ===== */
.calendar-area {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.calendar-toolbar {
  display: flex;
  align-items: center;
  padding: 8px 16px;
  gap: 8px;
  border-bottom: 1px solid var(--color-border);
  background: var(--color-surface);
  flex-shrink: 0;
}
.calendar-toolbar .nav-group {
  display: flex;
  align-items: center;
  gap: 4px;
}
.calendar-toolbar .date-range-label {
  font-size: 14px;
  font-weight: 600;
  color: var(--color-text);
  min-width: 200px;
  text-align: center;
}
.calendar-scroll-wrapper {
  flex: 1;
  overflow: auto;
  position: relative;
}
.calendar-scroll-wrapper::-webkit-scrollbar { width: 8px; height: 8px; }
.calendar-scroll-wrapper::-webkit-scrollbar-track { background: var(--color-bg-secondary); }
.calendar-scroll-wrapper::-webkit-scrollbar-thumb {
  background: var(--color-border-strong);
  border-radius: 4px;
}

/* Calendar Grid */
.calendar-grid {
  display: grid;
  position: relative;
  min-width: max-content;
}
.calendar-month-header {
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  min-width: max-content;
  background: var(--color-surface);
  border-bottom: 1px solid var(--color-border);
}
.calendar-month-cell {
  text-align: center;
  font-size: 13px;
  font-weight: 600;
  color: var(--color-text);
  padding: 4px 0;
  border-right: 1px solid var(--color-border);
  background: var(--color-bg-secondary);
}
.calendar-day-header {
  position: sticky;
  top: 28px;
  z-index: 9;
  display: flex;
  min-width: max-content;
  background: var(--color-surface);
  border-bottom: 1px solid var(--color-border-strong);
}
.calendar-day-cell {
  width: 40px;
  min-width: 40px;
  text-align: center;
  font-size: 11px;
  color: var(--color-text-secondary);
  padding: 4px 0;
  border-right: 1px solid var(--color-border);
  background: var(--color-surface);
}
.calendar-day-cell.weekend { background: var(--color-bg-tertiary); }
.calendar-day-cell.today {
  background: var(--color-primary) !important;
  color: var(--color-text-inverse) !important;
  font-weight: 700;
}

.calendar-body {
  position: relative;
  min-height: 200px;
}
.calendar-grid-lines {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  display: flex;
  pointer-events: none;
}
.calendar-grid-line {
  width: 40px;
  min-width: 40px;
  border-right: 1px solid var(--color-border);
  height: 100%;
}
.calendar-grid-line.weekend { background: var(--color-bg-tertiary); opacity: 0.3; }
.calendar-grid-line.today { background: var(--color-primary-light); opacity: 0.35; }

/* Calendar Blocks */
.calendar-block {
  position: absolute;
  border-radius: var(--radius-sm);
  padding: 4px 8px;
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  gap: 1px;
  overflow: hidden;
  cursor: pointer;
  transition: box-shadow var(--transition), transform var(--transition);
  font-size: 12px;
  border: 1px solid transparent;
}
.calendar-block.collapsed {
  max-height: var(--block-height, 90px);
}
.calendar-block.expanded {
  max-height: none;
  z-index: 15;
  box-shadow: var(--shadow-lg);
}
.calendar-block:hover {
  box-shadow: var(--shadow-md);
  transform: translateY(-1px);
  z-index: 5;
}
.calendar-block.expanded:hover { transform: none; }
.calendar-block.selected {
  border-color: var(--color-text);
  box-shadow: var(--shadow-lg);
  z-index: 6;
}
.calendar-block.expanded.selected { z-index: 16; }
.block-heading { font-weight: 600; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.block-subheading { font-size: 11px; opacity: 0.85; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.block-labels-row {
  display: flex;
  flex-direction: row;
  flex-wrap: wrap;
  gap: 4px;
  align-items: center;
}
.block-label {
  font-size: 10px;
  font-weight: 600;
  padding: 1px 6px;
  border-radius: 8px;
  background: rgba(255,255,255,0.3);
  white-space: nowrap;
}
.block-regular {
  font-size: 11px;
  opacity: 0.8;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.calendar-block.expanded .block-regular {
  -webkit-line-clamp: unset;
  line-clamp: unset;
  display: block;
}

/* Block color palette */
.block-color-0 { background: #0078d4; color: #fff; }
.block-color-1 { background: #00b7c3; color: #fff; }
.block-color-2 { background: #8764b8; color: #fff; }
.block-color-3 { background: #e3008c; color: #fff; }
.block-color-4 { background: #ca5010; color: #fff; }
.block-color-5 { background: #498205; color: #fff; }
.block-color-6 { background: #005b70; color: #fff; }
.block-color-7 { background: #8e562e; color: #fff; }

/* ===== Setup Wizard ===== */
.setup-overlay {
  position: fixed;
  inset: 0;
  background: var(--color-overlay);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 100;
}
.setup-panel {
  background: var(--color-surface);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-lg);
  width: 720px;
  max-width: 95vw;
  max-height: 90vh;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}
.setup-header {
  padding: 20px 24px 16px;
  border-bottom: 1px solid var(--color-border);
}
.setup-header h2 {
  font-size: 20px;
  font-weight: 600;
  color: var(--color-text);
}
.setup-header p {
  font-size: 13px;
  color: var(--color-text-secondary);
  margin-top: 4px;
}
.setup-steps {
  display: flex;
  gap: 4px;
  margin-top: 16px;
}
.setup-step-dot {
  width: 32px;
  height: 4px;
  border-radius: 2px;
  background: var(--color-border);
  transition: background var(--transition);
}
.setup-step-dot.active { background: var(--color-primary); }
.setup-step-dot.completed { background: var(--color-success); }
.setup-body {
  flex: 1;
  overflow-y: auto;
  padding: 24px;
}
.setup-footer {
  padding: 16px 24px;
  border-top: 1px solid var(--color-border);
  display: flex;
  justify-content: space-between;
  align-items: center;
}

/* Setup form elements */
.form-group {
  margin-bottom: 16px;
}
.form-group label {
  display: block;
  font-size: 13px;
  font-weight: 600;
  color: var(--color-text);
  margin-bottom: 4px;
}
.form-group .helper-text {
  font-size: 12px;
  color: var(--color-text-tertiary);
  margin-top: 2px;
}
.form-input {
  width: 100%;
  padding: 8px 12px;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-sm);
  font-family: var(--font-family);
  font-size: 14px;
  background: var(--color-bg);
  color: var(--color-text);
  transition: border-color var(--transition);
}
.form-input:focus {
  outline: none;
  border-color: var(--color-primary);
  box-shadow: 0 0 0 1px var(--color-primary);
}

/* Site browser */
.site-browser {
  border: 1px solid var(--color-border);
  border-radius: var(--radius-md);
  max-height: 250px;
  overflow-y: auto;
  margin-top: 8px;
}
.browser-item {
  display: flex;
  align-items: center;
  padding: 8px 12px;
  cursor: pointer;
  transition: background var(--transition);
  gap: 8px;
  border-bottom: 1px solid var(--color-border);
  font-size: 13px;
}
.browser-item:last-child { border-bottom: none; }
.browser-item:hover { background: var(--color-surface-hover); }
.browser-item.selected {
  background: var(--color-primary-light);
  color: var(--color-primary);
}
.browser-item .icon {
  font-size: 16px;
  flex-shrink: 0;
  width: 20px;
  text-align: center;
}
.browser-breadcrumb {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 8px 12px;
  font-size: 12px;
  color: var(--color-text-secondary);
  border-bottom: 1px solid var(--color-border);
  flex-wrap: wrap;
}
.browser-breadcrumb span {
  cursor: pointer;
  color: var(--color-primary);
}
.browser-breadcrumb span:hover { text-decoration: underline; }
.browser-breadcrumb .sep { color: var(--color-text-tertiary); cursor: default; }
.browser-loading {
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  color: var(--color-text-tertiary);
  font-size: 13px;
  gap: 8px;
}

/* Field mapping table */
.field-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
  table-layout: fixed;
}
.field-table th {
  text-align: left;
  padding: 6px 8px;
  font-weight: 600;
  color: var(--color-text-secondary);
  background: var(--color-bg-secondary);
  border-bottom: 1px solid var(--color-border);
  font-size: 12px;
  position: sticky;
  top: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.field-table th:nth-child(1) { width: 25%; }
.field-table th:nth-child(2) { width: 15%; }
.field-table th:nth-child(3) { width: 20%; }
.field-table th:nth-child(4) { width: 25%; }
.field-table th:nth-child(5) { width: 15%; }
.field-table td {
  padding: 5px 8px;
  border-bottom: 1px solid var(--color-border);
  vertical-align: middle;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.field-table tr:hover td { background: var(--color-surface-hover); }
.field-table select {
  font-family: var(--font-family);
  font-size: 12px;
  padding: 3px 4px;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-sm);
  background: var(--color-bg);
  color: var(--color-text);
  width: 100%;
  max-width: 100%;
}
.sample-value {
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--color-text-tertiary);
  font-size: 12px;
}

/* ===== Modal ===== */
.modal-overlay {
  position: fixed;
  inset: 0;
  background: var(--color-overlay);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 200;
}
.modal-overlay.hidden { display: none; }
.modal-panel {
  background: var(--color-surface);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-lg);
  width: 520px;
  max-width: 95vw;
  max-height: 80vh;
  overflow: auto;
  padding: 24px;
}
.modal-panel h3 {
  font-size: 18px;
  font-weight: 600;
  margin-bottom: 12px;
}
.modal-panel .code-block {
  background: var(--color-bg-tertiary);
  border: 1px solid var(--color-border);
  border-radius: var(--radius-sm);
  padding: 12px;
  font-family: 'Consolas', 'Courier New', monospace;
  font-size: 12px;
  color: var(--color-text);
  word-break: break-all;
  margin: 12px 0;
  position: relative;
}
.modal-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 16px;
}

/* ===== Permission Banner ===== */
.permission-banner {
  background: var(--color-warning);
  color: #323130;
  padding: 10px 20px;
  font-size: 13px;
  font-weight: 600;
  text-align: center;
  flex-shrink: 0;
}

/* ===== Error/Info messages ===== */
.msg-error {
  background: #fde7e9;
  color: var(--color-danger);
  padding: 8px 12px;
  border-radius: var(--radius-sm);
  font-size: 13px;
  margin-top: 8px;
}
.msg-info {
  background: var(--color-primary-light);
  color: var(--color-primary);
  padding: 8px 12px;
  border-radius: var(--radius-sm);
  font-size: 13px;
  margin-top: 8px;
}

/* ===== Empty state ===== */
.empty-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 60px 20px;
  text-align: center;
  color: var(--color-text-tertiary);
}
.empty-state .icon { font-size: 48px; margin-bottom: 12px; }
.empty-state h3 { font-size: 16px; font-weight: 600; color: var(--color-text-secondary); }
.empty-state p { font-size: 13px; margin-top: 4px; }

/* ===== Create fields prompt ===== */
.create-fields-prompt {
  background: var(--color-primary-light);
  border: 1px solid var(--color-primary);
  border-radius: var(--radius-md);
  padding: 16px;
  margin-top: 16px;
}
.create-fields-prompt h4 {
  font-size: 14px;
  font-weight: 600;
  color: var(--color-primary);
  margin-bottom: 8px;
}
.create-fields-prompt p {
  font-size: 13px;
  color: var(--color-text-secondary);
  margin-bottom: 12px;
}

/* ===== Mobile View Switcher (Desktop hidden) ===== */
.mobile-view-switcher {
  display: none;
}

/* ===== Responsive ===== */
@media (max-width: 768px) {
  .header-actions {
    display: none !important;
  }

  .main-container {
    flex-direction: column;
    overflow: hidden;
  }
  
  /* Mobile View Switcher */
  .mobile-view-switcher {
    display: flex;
    justify-content: center;
    background: var(--color-surface);
    border-bottom: 1px solid var(--color-border);
    padding: 8px 16px;
    gap: 12px;
    flex-shrink: 0;
    width: 100%;
    align-items: center;
  }
  .mobile-view-switcher .switch-btn {
    flex: 1;
    max-width: 160px;
    padding: 8px 16px;
    font-size: 13px;
    font-weight: 600;
    border: 1px solid var(--color-border);
    background: var(--color-bg);
    color: var(--color-text-secondary);
    border-radius: var(--radius-md);
    cursor: pointer;
    transition: all 0.2s ease;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    outline: none;
  }
  .mobile-view-switcher .switch-btn:hover {
    background: var(--color-surface-hover);
    color: var(--color-text);
  }
  .mobile-view-switcher .switch-btn:active {
    transform: scale(0.97);
  }
  .mobile-view-switcher .switch-btn.active {
    background: var(--color-primary);
    color: #ffffff;
    border-color: var(--color-primary);
    box-shadow: 0 2px 4px rgba(0, 120, 212, 0.15);
  }
  .mobile-view-switcher .switch-btn-icon {
    font-size: 14px;
    line-height: 1;
  }

  /* Responsive View Toggles */
  .main-container.mobile-calendar .sidebar {
    display: none !important;
  }
  .main-container.mobile-calendar .calendar-area {
    display: flex !important;
    flex: 1;
    overflow: hidden;
  }

  .main-container.mobile-list .sidebar {
    display: flex !important;
    flex: 1;
    width: 100%;
    max-width: 100%;
    min-width: 0;
    border-right: none;
    border-bottom: none;
    max-height: none !important;
    height: 100%;
    overflow: hidden;
  }
  .main-container.mobile-list .calendar-area {
    display: none !important;
  }
}

/* ===== Month Grid View (Week Rows) ===== */
.calendar-scroll-wrapper.view-month {
  overflow-x: hidden;
  overflow-y: auto;
}
.calendar-scroll-wrapper.view-month #calendarMonthHeader {
  display: none !important;
}
.calendar-scroll-wrapper.view-month #calendarDayHeader {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  width: 100%;
  min-width: 600px;
  position: sticky;
  top: 0;
  z-index: 10;
  border-bottom: 1px solid var(--color-border-strong);
}
.calendar-scroll-wrapper.view-month .calendar-day-name-cell {
  padding: 8px 4px;
  text-align: center;
  font-size: 12px;
  font-weight: 600;
  color: var(--color-text-secondary);
  background: var(--color-bg-secondary);
  border-right: 1px solid var(--color-border);
}
.calendar-scroll-wrapper.view-month .calendar-day-name-cell:last-child {
  border-right: none;
}
.calendar-scroll-wrapper.view-month .calendar-day-name-cell.weekend {
  background: var(--color-bg-tertiary);
}

.calendar-scroll-wrapper.view-month #calendarBody {
  display: flex;
  flex-direction: column;
  width: 100%;
  min-width: 600px;
  min-height: 100%;
  position: relative;
}
.calendar-scroll-wrapper.view-month #calendarGridLines {
  display: none !important;
}

/* Week Row container */
.week-row {
  display: flex;
  flex-direction: column;
  position: relative;
  border-bottom: 1px solid var(--color-border);
  min-height: 120px;
  background: var(--color-surface);
  flex: 1;
}
.week-row:last-child {
  border-bottom: none;
}

/* 7-day Background cell grid */
.week-grid-days {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  position: absolute;
  inset: 0;
  pointer-events: none;
}
.week-day-cell {
  border-right: 1px solid var(--color-border);
  height: 100%;
  position: relative;
  pointer-events: auto;
}
.week-day-cell:last-child {
  border-right: none;
}
.week-day-cell.weekend {
  background: var(--color-bg-secondary);
  opacity: 0.5;
}
.week-day-cell.today {
  background: var(--color-primary-light);
  box-shadow: inset 0 0 0 2px var(--color-primary);
}
.week-day-cell.outside-month {
  background: var(--color-bg-secondary);
  opacity: 0.35;
}

/* Day Number Label */
.day-number-container {
  position: absolute;
  top: 4px;
  right: 6px;
  display: flex;
  align-items: center;
  justify-content: center;
  pointer-events: none;
}
.day-number-label {
  font-size: 11px;
  font-weight: 600;
  color: var(--color-text-secondary);
  width: 22px;
  height: 22px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 50%;
}
.week-day-cell.today .day-number-label {
  background: var(--color-primary);
  color: var(--color-text-inverse);
  box-shadow: var(--shadow-sm);
}
.week-day-cell.outside-month .day-number-label {
  color: var(--color-text-tertiary);
}

/* Week Events stack container */
.week-events-container {
  margin-top: 28px;
  margin-bottom: 4px;
  position: relative;
  flex-grow: 1;
}
.week-events-container .calendar-block {
  position: absolute;
  box-sizing: border-box;
}
</style>
</head>
<body>
<!-- Permission Banner (shown when user lacks Contribute permissions) -->
<div id="permissionBanner" class="permission-banner" style="display:none;">
  You have read-only access. A user with Contribute permissions is required to make changes or complete setup.
</div>

<!-- Limit Exhausted Banner (shown when the loaded item count matches the custom fetching limit) -->
<div id="limitBanner" class="permission-banner" style="display:none; background:#fff2cc; color:#856404; border-bottom: 1px solid #ffeeba; position:relative; padding-right: 40px;">
  &#9888; Item limit reached. Some older or unmodified items may be omitted. If you need to view more items, please increase the "Maximum Items to Fetch" limit in the <a href="#" id="linkReconfigureLimit" style="color:#856404; font-weight:700; text-decoration:underline;">Setup Configuration</a>.
  <button id="btnCloseLimitBanner" style="position:absolute; right:12px; top:50%; transform:translateY(-50%); border:none; background:transparent; font-size:16px; font-weight:700; color:#856404; cursor:pointer; padding:4px 8px; line-height:1;" title="Dismiss for 2 weeks">&times;</button>
</div>

<!-- Page Header -->
<div id="pageHeader" class="page-header">
  <div id="pageTitle" class="page-title">SharePoint Calendar</div>
  <div class="header-actions">
    <div id="saveIndicator" class="save-indicator">
      <span id="saveSpinner" class="spinner" style="display:none;"></span>
      <span id="saveCheck" class="checkmark" style="display:none;">&#10003;</span>
      <span id="saveText"></span>
    </div>
    <button id="btnReconfigure" class="btn btn-ghost btn-sm" style="display:none;" title="Reconfigure">&#9881; Setup</button>
    <button id="btnEmbed" class="btn btn-secondary btn-sm" style="display:none;">&#60;&#47;&#62; Embed</button>
  </div>
</div>

<!-- Main Layout -->
<div id="mainContainer" class="main-container mobile-calendar" style="display:none;">
  <!-- Mobile View Switcher (visible on mobile only) -->
  <div id="mobileViewSwitcher" class="mobile-view-switcher">
    <button id="btnViewCalendar" class="switch-btn active">
      <span class="switch-btn-icon">&#128197;</span> Calendar
    </button>
    <button id="btnViewList" class="switch-btn">
      <span class="switch-btn-icon">&#128203;</span> List
    </button>
  </div>

  <!-- Left Sidebar: Card Panel -->
  <div class="sidebar">
    <div class="sidebar-header">
      <span id="sidebarTitle">All Items</span>
      <span id="sidebarCount" style="font-weight:400; font-size:12px; color:var(--color-text-tertiary);"></span>
    </div>
    <div id="sidebarTabs" class="sidebar-tabs"></div>
    <div id="sidebarScroll" class="sidebar-scroll"></div>
  </div>
  <!-- Calendar Area -->
  <div class="calendar-area">
    <div class="calendar-toolbar">
      <div class="nav-group">
        <button id="btnPrevMonth" class="btn-icon" title="Previous month">&#9664;</button>
        <button id="btnToday" class="btn btn-ghost btn-sm">Today</button>
        <button id="btnNextMonth" class="btn-icon" title="Next month">&#9654;</button>
      </div>
      <div id="dateRangeLabel" class="date-range-label"></div>
      <div class="view-group" style="margin-left: auto; display: flex; gap: 4px;">
        <button id="btnViewMonthGrid" class="btn btn-ghost btn-sm" title="Month Grid View">Month</button>
        <button id="btnViewTimeline" class="btn btn-ghost btn-sm" title="Timeline Gantt View">Timeline</button>
      </div>
    </div>
    <div id="calendarScrollWrapper" class="calendar-scroll-wrapper">
      <div id="calendarMonthHeader" class="calendar-month-header"></div>
      <div id="calendarDayHeader" class="calendar-day-header"></div>
      <div id="calendarBody" class="calendar-body">
        <div id="calendarGridLines" class="calendar-grid-lines"></div>
      </div>
    </div>
  </div>
</div>

<!-- Embed Modal -->
<div id="embedModal" class="modal-overlay hidden">
  <div class="modal-panel">
    <h3>&#60;&#47;&#62; Embed this Calendar</h3>
    <p style="font-size:13px; color:var(--color-text-secondary);">
      Copy the code below and paste it into the <strong>Embed web part</strong> on a SharePoint Modern page.
    </p>
    <div id="embedCode" class="code-block"></div>
    <div class="modal-actions">
      <button id="btnCopyEmbed" class="btn btn-primary btn-sm">Copy to Clipboard</button>
      <button id="btnCloseEmbed" class="btn btn-ghost btn-sm">Close</button>
    </div>
  </div>
</div>

<script>
/* ===========================================================================
   SPCalendar - Single-Page SharePoint Calendar Application
   ===========================================================================
   All application logic is contained in this single script block.
   Uses SharePoint REST API exclusively. No external dependencies.
   =========================================================================== */
(function () {
  'use strict';

  // ===== State =====
  var savedView = localStorage.getItem('SPCalendar_PreferredView');
  var APP = {
    config: null,          // { listUrl, siteUrl, listTitle, listId, fieldMappings, startDateField, endDateField, pageTitle, entityType }
    items: [],             // SharePoint list items
    csvRows: [],           // Parsed CSV rows
    currentUser: null,     // { Title, Email }
    hasContribute: true,   // Permission flag
    isEmbedMode: false,
    calendarStartDate: null,
    calendarEndDate: null,
    selectedItemId: null,
    pollTimerCSV: null,
    pollTimerList: null,
    lastCSVLength: 0,
    saving: false,
    addingDatesForId: null, // Item ID currently in "add to calendar" date pick mode
    currentView: savedView === 'timeline' ? 'timeline' : 'month' // Default is 'month'
  };

  var DAY_MS = 86400000;
  var CELL_WIDTH = 40;
  var BLOCK_HEIGHT = 90;
  var BLOCK_GAP = 4;
  var BLOCK_TOP_OFFSET = 4;
  var STYLE_ORDER = { heading: 0, subheading: 1, regular: 2, label: 3 };

  // Locale-aware helpers
  function getLocaleFirstDayOfWeek() {
    try {
      var locale = navigator.language || 'en-US';
      if (typeof Intl !== 'undefined' && Intl.Locale) {
        var loc = new Intl.Locale(locale);
        var firstDay = null;
        if (loc.weekInfo && typeof loc.weekInfo.firstDay === 'number') {
          firstDay = loc.weekInfo.firstDay;
        } else if (typeof loc.getWeekInfo === 'function') {
          var info = loc.getWeekInfo();
          if (info && typeof info.firstDay === 'number') {
            firstDay = info.firstDay;
          }
        }
        if (firstDay !== null) {
          return firstDay === 7 ? 0 : firstDay;
        }
      }
    } catch (e) {
      console.warn('Error reading Intl.Locale weekInfo:', e);
    }

    var lang = (navigator.language || 'en-US').toLowerCase();
    var country = lang.split('-')[1] || lang;
    var sundayCountries = ['us', 'ca', 'mx', 'jp', 'br', 'in', 'kr', 'au', 'nz', 'il', 'za', 'ph', 'sg', 'hk'];
    if (sundayCountries.indexOf(country) !== -1) {
      return 0; // Sunday
    }
    return 1; // Monday (default)
  }

  function getLocaleWeekdayNames() {
    try {
      var locale = navigator.language || 'en-US';
      var formatter = new Intl.DateTimeFormat(locale, { weekday: 'short', timeZone: 'UTC' });
      var names = [];
      // Day 0 in UTC for standard date: 2026-05-24 is Sunday
      for (var i = 0; i < 7; i++) {
        var d = new Date(Date.UTC(2026, 4, 24 + i));
        names.push(formatter.format(d));
      }
      return names;
    } catch (e) {
      return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    }
  }

  // Central function to format any field value for display
  function formatFieldValue(val) {
    if (val == null || val === '') return '';
    if (typeof val === 'object' && val !== null) {
      // Person/User field (expanded): has Title
      if (val.Title) return val.Title;
      // Multi-value array (MultiChoice, MultiUser, MultiLookup)
      if (val.results && Array.isArray(val.results)) {
        return val.results.map(function (r) {
          if (typeof r === 'object') return r.Title || r.LookupValue || String(r);
          return String(r);
        }).join(', ');
      }
      // Deferred lookup (not expanded)
      if (val.__deferred) return '';
      // Fallback
      try { return JSON.stringify(val); } catch (e) { return ''; }
    }
    return String(val);
  }

  function sortedMappings() {
    return APP.config.fieldMappings.slice().sort(function (a, b) {
      var aOrder = STYLE_ORDER[a.style] != null ? STYLE_ORDER[a.style] : 99;
      var bOrder = STYLE_ORDER[b.style] != null ? STYLE_ORDER[b.style] : 99;
      return aOrder - bOrder;
    });
  }

  // ===== DOM Helpers =====
  function el(id) { return document.getElementById(id); }

  function create(tag, attrs, children) {
    var node = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        if (k === 'className') node.className = attrs[k];
        else if (k === 'textContent') node.textContent = attrs[k];
        else if (k.indexOf('on') === 0) node.addEventListener(k.slice(2).toLowerCase(), attrs[k]);
        else node.setAttribute(k, attrs[k]);
      });
    }
    if (children) {
      children.forEach(function (c) {
        if (typeof c === 'string') node.appendChild(document.createTextNode(c));
        else if (c) node.appendChild(c);
      });
    }
    return node;
  }

  function clearEl(node) { node.replaceChildren(); }

  function show(node) { node.style.display = ''; }
  function hide(node) { node.style.display = 'none'; }

  // ===== Date Helpers =====
  function toDateStr(d) {
    if (!d) return '';
    var date = typeof d === 'string' ? new Date(d) : d;
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    var y = date.getFullYear();
    var m = String(date.getMonth() + 1).padStart(2, '0');
    var dd = String(date.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + dd;
  }

  function parseDate(s) {
    if (!s) return null;
    var d = new Date(s);
    if (isNaN(d.getTime())) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function localDateToISO(dateStr) {
    if (!dateStr) return null;
    var parts = dateStr.split('-');
    if (parts.length !== 3) return null;
    var y = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10) - 1;
    var d = parseInt(parts[2], 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
    // Construct local date at noon (12:00:00) to safely avoid timezone offset shift bugs
    var date = new Date(y, m, d, 12, 0, 0);
    if (isNaN(date.getTime())) return null;
    return date.toISOString();
  }


  function daysBetween(a, b) {
    return Math.round((b.getTime() - a.getTime()) / DAY_MS);
  }

  function addDays(d, n) {
    var r = new Date(d);
    r.setDate(r.getDate() + n);
    return r;
  }

  function isWeekend(d) { var day = d.getDay(); return day === 0 || day === 6; }

  function isSameDay(a, b) {
    return a.getFullYear() === b.getFullYear() &&
           a.getMonth() === b.getMonth() &&
           a.getDate() === b.getDate();
  }

  var MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var MONTH_FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];

  // ===== SharePoint REST Helpers =====
  var _digestValue = null;
  var _digestExpiry = 0;

  function getSiteUrl() {
    // When running within SharePoint, try _spPageContextInfo first
    if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo.webAbsoluteUrl) {
      return _spPageContextInfo.webAbsoluteUrl;
    }
    // Fallback: derive from current page URL
    var url = window.location.href;
    var match = url.match(/(https?:\/\/[^/]+\/sites\/[^/]+)/i) ||
                url.match(/(https?:\/\/[^/]+)/i);
    return match ? match[1] : '';
  }

  function getPageLibraryUrl() {
    // Get the folder URL where this .aspx lives
    var path = window.location.pathname;
    var lastSlash = path.lastIndexOf('/');
    return path.substring(0, lastSlash);
  }

  function getRequestDigest() {
    return new Promise(function (resolve, reject) {
      // Check cached
      if (_digestValue && Date.now() < _digestExpiry) {
        resolve(_digestValue);
        return;
      }
      // Try page element (classic pages)
      var digestEl = document.getElementById('__REQUESTDIGEST');
      if (digestEl && digestEl.value) {
        _digestValue = digestEl.value;
        _digestExpiry = Date.now() + 1500000; // ~25 min
        resolve(_digestValue);
        return;
      }
      // Fetch via REST
      var siteUrl = getSiteUrl();
      fetch(siteUrl + '/_api/contextinfo', {
        method: 'POST',
        headers: { 'Accept': 'application/json;odata=verbose' },
        credentials: 'include'
      })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        _digestValue = data.d.GetContextWebInformation.FormDigestValue;
        _digestExpiry = Date.now() + 1500000;
        resolve(_digestValue);
      })
      .catch(reject);
    });
  }

  function spGet(url) {
    return fetch(url, {
      method: 'GET',
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'include'
    }).then(function (r) {
      if (!r.ok) throw new Error('GET failed: ' + r.status + ' ' + url);
      return r.json();
    });
  }

  function spPost(url, body, extraHeaders) {
    return getRequestDigest().then(function (digest) {
      var headers = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      };
      if (extraHeaders) Object.assign(headers, extraHeaders);
      return fetch(url, {
        method: 'POST',
        headers: headers,
        credentials: 'include',
        body: typeof body === 'string' ? body : JSON.stringify(body)
      });
    }).then(function (r) {
      if (!r.ok) throw new Error('POST failed: ' + r.status + ' ' + url);
      var ct = r.headers.get('content-type') || '';
      if (ct.indexOf('json') !== -1) return r.json();
      return r.text();
    });
  }

  function spMerge(url, body, entityType) {
    return getRequestDigest().then(function (digest) {
      return fetch(url, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        credentials: 'include',
        body: JSON.stringify(Object.assign({ '__metadata': { 'type': entityType } }, body))
      });
    }).then(function (r) {
      if (!r.ok) throw new Error('MERGE failed: ' + r.status);
    });
  }

  // ===== User & Permissions =====
  function getCurrentUser() {
    var siteUrl = getSiteUrl();
    return spGet(siteUrl + "/_api/web/currentuser?$select=Title,Email").then(function (data) {
      var d = data && data.d ? data.d : {};
      APP.currentUser = { Title: d.Title || 'Unknown User', Email: d.Email || '' };
      return APP.currentUser;
    });
  }

  function checkPermissions() {
    var siteUrl = getSiteUrl();
    // Primary: check web-level permissions (most reliable)
    return spGet(siteUrl + "/_api/web/effectivebasepermissions")
      .then(function (data) {
        var d = data && data.d ? data.d : {};
        var perms = d.EffectiveBasePermissions || d;
        var low = perms.Low || 0;
        // AddListItems permission = bit 1 in the Low word
        APP.hasContribute = (low & 2) !== 0;
      })
      .catch(function () {
        // Fallback: try folder-level check on the library where this file lives
        var libPath = getPageLibraryUrl();
        return spGet(siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(libPath) + "')/ListItemAllFields/EffectiveBasePermissions")
          .then(function (data) {
            var d = data && data.d ? data.d : {};
            var perms = d.EffectiveBasePermissions || d;
            var low = perms.Low || 0;
            APP.hasContribute = (low & 2) !== 0;
          })
          .catch(function () {
            // Final fallback: practical test - if we can get a request digest,
            // the user can make write requests
            return fetch(siteUrl + '/_api/contextinfo', {
              method: 'POST',
              headers: { 'Accept': 'application/json;odata=verbose' },
              credentials: 'include'
            }).then(function (r) {
              APP.hasContribute = r.ok;
            }).catch(function () {
              APP.hasContribute = false;
            });
          });
      });
  }

  // ===== CSV State Store =====
  var CSV_FILENAME = 'SPCalendar_ActionLog.csv';
  var CSV_HEADERS = 'Timestamp,UserName,UserEmail,ActionType,ItemID,Key,Value,Details';

  function csvEscape(val) {
    if (val == null) return '';
    var s = String(val);
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function csvRowToString(row) {
    return [
      csvEscape(row.Timestamp),
      csvEscape(row.UserName),
      csvEscape(row.UserEmail),
      csvEscape(row.ActionType),
      csvEscape(row.ItemID),
      csvEscape(row.Key),
      csvEscape(row.Value),
      csvEscape(row.Details)
    ].join(',');
  }

  function parseCSVLine(line) {
    var result = [];
    var current = '';
    var inQuotes = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (inQuotes) {
        if (c === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += c;
        }
      } else {
        if (c === '"') {
          inQuotes = true;
        } else if (c === ',') {
          result.push(current);
          current = '';
        } else {
          current += c;
        }
      }
    }
    result.push(current);
    return result;
  }

  function parseCSV(text) {
    if (!text || !text.trim()) return [];
    var lines = text.split('\n').filter(function (l) { return l.trim(); });
    if (lines.length <= 1) return []; // Only header
    var rows = [];
    for (var i = 1; i < lines.length; i++) {
      var fields = parseCSVLine(lines[i]);
      if (fields.length >= 8) {
        rows.push({
          Timestamp: fields[0],
          UserName: fields[1],
          UserEmail: fields[2],
          ActionType: fields[3],
          ItemID: fields[4],
          Key: fields[5],
          Value: fields[6],
          Details: fields[7]
        });
      }
    }
    return rows;
  }

  function csvToString(rows) {
    var lines = [CSV_HEADERS];
    rows.forEach(function (r) { lines.push(csvRowToString(r)); });
    return lines.join('\n');
  }

  function loadCSV() {
    var siteUrl = getSiteUrl();
    var libPath = getPageLibraryUrl();
    var fileUrl = siteUrl + "/_api/web/GetFileByServerRelativeUrl('" +
      encodeURIComponent(libPath + '/' + CSV_FILENAME) + "')/$value";
    return fetch(fileUrl, {
      method: 'GET',
      credentials: 'include'
    }).then(function (r) {
      if (r.status === 404) return '';
      if (!r.ok) throw new Error('CSV load failed: ' + r.status);
      return r.text();
    }).then(function (text) {
      APP.csvRows = parseCSV(text);
      APP.lastCSVLength = APP.csvRows.length;
      return APP.csvRows;
    }).catch(function () {
      APP.csvRows = [];
      APP.lastCSVLength = 0;
      return [];
    });
  }

  function saveCSV() {
    var siteUrl = getSiteUrl();
    var libPath = getPageLibraryUrl();
    var csvText = csvToString(APP.csvRows);
    var uploadUrl = siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" +
      encodeURIComponent(libPath) + "')/Files/add(url='" + CSV_FILENAME + "',overwrite=true)";
    return getRequestDigest().then(function (digest) {
      return fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'X-RequestDigest': digest
        },
        credentials: 'include',
        body: csvText
      });
    }).then(function (r) {
      if (!r.ok) throw new Error('CSV save failed: ' + r.status);
      APP.lastCSVLength = APP.csvRows.length;
    });
  }

  function appendCSVRow(actionType, itemId, key, value, details) {
    var row = {
      Timestamp: new Date().toISOString(),
      UserName: APP.currentUser ? APP.currentUser.Title : '',
      UserEmail: APP.currentUser ? APP.currentUser.Email : '',
      ActionType: actionType,
      ItemID: itemId || '',
      Key: key || '',
      Value: value || '',
      Details: details || ''
    };
    APP.csvRows.push(row);
    return row;
  }

  function replayConfig() {
    // Replay SETUP_CONFIG rows to reconstruct APP.config
    var cfg = {};
    var archived = {};
    APP.csvRows.forEach(function (row) {
      if (row.ActionType === 'SETUP_CONFIG') {
        try {
          cfg[row.Key] = JSON.parse(row.Value);
        } catch (e) {
          cfg[row.Key] = row.Value;
        }
      } else if (row.ActionType === 'ARCHIVE_ITEM') {
        if (row.Value === 'true') {
          archived[row.ItemID] = true;
        } else {
          delete archived[row.ItemID];
        }
      }
    });
    APP._archivedIds = archived;
    if (cfg.listUrl && cfg.fieldMappings) {
      APP.config = {
        listUrl: cfg.listUrl,
        siteUrl: cfg.siteUrl || getSiteUrl(),
        listTitle: cfg.listTitle || '',
        listId: cfg.listId || '',
        fieldMappings: cfg.fieldMappings,
        startDateField: cfg.startDateField || null,
        endDateField: cfg.endDateField || null,
        pageTitle: cfg.pageTitle || 'SharePoint Calendar',
        entityType: cfg.entityType || '',
        itemLimit: parseInt(cfg.itemLimit, 10) || 5000,
        colorBaseField: cfg.colorBaseField || null,
        colorBaseFieldType: cfg.colorBaseFieldType || null,
        colorMap: cfg.colorMap || null
      };
      return true;
    }
    return false;
  }

  // ===== Save Indicator =====
  function showSaving() {
    APP.saving = true;
    var ind = el('saveIndicator');
    show(el('saveSpinner'));
    hide(el('saveCheck'));
    el('saveText').textContent = 'Saving...';
    ind.classList.add('visible');
  }

  function showSaved() {
    APP.saving = false;
    hide(el('saveSpinner'));
    show(el('saveCheck'));
    el('saveText').textContent = 'Changes Saved!';
    setTimeout(function () {
      if (!APP.saving) {
        el('saveIndicator').classList.remove('visible');
      }
    }, 3000);
  }

  function showSaveError() {
    APP.saving = false;
    hide(el('saveSpinner'));
    hide(el('saveCheck'));
    el('saveText').textContent = 'Save failed';
    el('saveIndicator').classList.add('visible');
    setTimeout(function () {
      if (!APP.saving) el('saveIndicator').classList.remove('visible');
    }, 5000);
  }

  // ===== Setup Wizard =====
  var setupStep = 0;
  var setupData = {
    siteUrl: '',
    listTitle: '',
    listId: '',
    listUrl: '',
    fields: [],
    sampleItem: null,
    fieldMappings: [],
    startDateField: null,
    endDateField: null,
    pageTitle: '',
    entityType: '',
    hasDateFields: false,
    itemLimit: 5000,
    colorBaseField: null,
    colorBaseFieldType: null,
    colorMap: null,
    uniqueValues: []
  };

  var PRESET_COLORS = [
    '#0078D4', // Royal Blue
    '#D13438', // Crimson Red
    '#107C10', // Forest Green
    '#FF8F00', // Warm Amber
    '#8764B8', // Royal Purple
    '#CA5010', // Vibrant Orange
    '#00B7C3', // Bright Teal
    '#E3008C', // Deep Pink
    '#8E562E', // Warm Brown
    '#004B87', // Navy Blue
    '#498205', // Olive Green
    '#881794', // Rich Violet
    '#D69E2E', // Warm Gold
    '#005B70', // Dark Teal
    '#9E1C3F', // Crimson Rose
    '#605E5C'  // Slate Gray
  ];

  function getUniqueValuesForField(items, fieldName) {
    var uniqueSet = {};
    items.forEach(function (item) {
      var val = item[fieldName];
      if (val == null || val === '') return;
      var strVal = formatFieldValue(val);
      if (strVal && strVal.trim()) {
        uniqueSet[strVal.trim()] = true;
      }
    });
    return Object.keys(uniqueSet);
  }

  function getFieldTitle(fieldName) {
    var field = setupData.fields.find(function (f) { return f.InternalName === fieldName; });
    return field ? field.Title : fieldName;
  }

  function showSetup() {
    setupStep = 0;
    setupData.siteUrl = getSiteUrl();

    // If reconfiguring, pre-populate setupData from existing config
    var isReconfigure = !!APP.config;
    if (isReconfigure) {
      setupData.siteUrl = APP.config.siteUrl || getSiteUrl();
      setupData.listTitle = APP.config.listTitle || '';
      setupData.listId = APP.config.listId || '';
      setupData.listUrl = APP.config.listUrl || '';
      setupData.fieldMappings = APP.config.fieldMappings ? APP.config.fieldMappings.slice() : [];
      setupData.startDateField = APP.config.startDateField || null;
      setupData.endDateField = APP.config.endDateField || null;
      setupData.pageTitle = APP.config.pageTitle || '';
      setupData.entityType = APP.config.entityType || '';
      setupData.itemLimit = APP.config.itemLimit || 5000;
      setupData.colorBaseField = APP.config.colorBaseField || null;
      setupData.colorBaseFieldType = APP.config.colorBaseFieldType || null;
      setupData.colorMap = APP.config.colorMap ? Object.assign({}, APP.config.colorMap) : null;
    } else {
      setupData.colorBaseField = null;
      setupData.colorBaseFieldType = null;
      setupData.colorMap = null;
      setupData.uniqueValues = [];
    }

    var overlay = create('div', { className: 'setup-overlay', id: 'setupOverlay' });
    var panel = create('div', { className: 'setup-panel' });

    var header = create('div', { className: 'setup-header', style: 'position:relative;' });
    var h2 = create('h2', { id: 'setupTitle', textContent: 'Setup Calendar' });
    var p = create('p', { id: 'setupSubtitle', textContent: 'Connect to a SharePoint list to get started.' });
    var dots = create('div', { className: 'setup-steps', id: 'setupDots' });
    for (var i = 0; i < 4; i++) {
      dots.appendChild(create('div', { className: 'setup-step-dot' + (i === 0 ? ' active' : ''), 'data-step': String(i) }));
    }
    header.appendChild(h2);
    header.appendChild(p);
    header.appendChild(dots);

    // Close button (only on reconfigure, not first-time setup)
    if (isReconfigure) {
      var closeBtn = create('button', {
        className: 'btn-icon',
        style: 'position:absolute;top:8px;right:8px;font-size:18px;width:28px;height:28px;display:flex;align-items:center;justify-content:center;',
        textContent: '\u2715',
        title: 'Close',
        onClick: function () {
          var ov = el('setupOverlay');
          if (ov) ov.remove();
          show(el('mainContainer'));
        }
      });
      header.appendChild(closeBtn);
    }

    var body = create('div', { className: 'setup-body', id: 'setupBody' });
    var footer = create('div', { className: 'setup-footer', id: 'setupFooter' });

    panel.appendChild(header);
    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    // If reconfiguring and we have list info, jump to field mapping with preloaded data
    if (isReconfigure && setupData.listUrl && setupData.fields && setupData.fields.length > 0) {
      setupStep = 1;
    }

    renderSetupStep();
  }

  function renderSetupStep() {
    var body = el('setupBody');
    var footer = el('setupFooter');
    clearEl(body);
    clearEl(footer);

    // Update dots
    var dots = el('setupDots').children;
    for (var i = 0; i < dots.length; i++) {
      dots[i].className = 'setup-step-dot';
      if (i < setupStep) dots[i].classList.add('completed');
      if (i === setupStep) dots[i].classList.add('active');
    }
    // If color base is not selected and we are on step 3, mark step 2 dot as completed
    if (setupStep === 3 && !setupData.colorBaseField) {
      if (dots[2]) dots[2].classList.add('completed');
    }

    if (setupStep === 0) renderSetupStep1(body, footer);
    else if (setupStep === 1) renderSetupStep2(body, footer);
    else if (setupStep === 2) renderSetupStep2ColorConfig(body, footer);
    else if (setupStep === 3) renderSetupStep3(body, footer);
  }

  // ----- Step 1: Connect to List -----
  function renderSetupStep1(body, footer) {
    el('setupTitle').textContent = 'Connect to SharePoint List';
    el('setupSubtitle').textContent = 'Select a list from your site, or enter a URL manually.';

    if (!APP.hasContribute) {
      body.appendChild(create('div', { className: 'msg-error', textContent: 'A user with Contribute permissions is required to complete setup. Please contact your SharePoint administrator.' }));
      return;
    }

    // Site browser (primary)
    var browserLabel = create('label', { textContent: 'Select a list:', style: 'font-size:13px;font-weight:600;display:block;margin-bottom:8px;' });
    body.appendChild(browserLabel);

    var browserContainer = create('div', { className: 'site-browser', id: 'siteBrowser' });
    body.appendChild(browserContainer);

    loadSiteBrowser(browserContainer, setupData.siteUrl);

    // URL input (fallback)
    var grp = create('div', { className: 'form-group', style: 'margin-top:16px;' });
    grp.appendChild(create('label', { textContent: 'Or enter a list URL manually', for: 'inputListUrl' }));
    var inp = create('input', {
      className: 'form-input',
      id: 'inputListUrl',
      type: 'text',
      placeholder: 'e.g. https://contoso.sharepoint.com/sites/MySite/Lists/Projects'
    });
    if (setupData.listUrl) inp.value = setupData.listUrl;
    grp.appendChild(inp);
    body.appendChild(grp);

    // Error area
    body.appendChild(create('div', { id: 'step1Error' }));

    // Footer
    var btnNext = create('button', {
      className: 'btn btn-primary',
      textContent: 'Next \u2192',
      onClick: function () { validateStep1(); }
    });
    footer.appendChild(create('span'));
    footer.appendChild(btnNext);
  }

  function loadSiteBrowser(container, siteUrl) {
    clearEl(container);
    container.appendChild(create('div', { className: 'browser-loading' }, [
      create('div', { className: 'spinner', style: 'width:14px;height:14px;' }),
      'Loading...'
    ]));

    var breadcrumb = create('div', { className: 'browser-breadcrumb' });
    var itemsContainer = create('div');

    // Load subsites and lists
    Promise.all([
      spGet(siteUrl + "/_api/web/webinfos?$select=ServerRelativeUrl,Title").catch(function () { return { d: { results: [] } }; }),
      spGet(siteUrl + "/_api/web/lists?$filter=Hidden eq false&$select=Title,Id,BaseTemplate,ItemCount&$orderby=Title").catch(function () { return { d: { results: [] } }; }),
      spGet(siteUrl + "/_api/web?$select=Title,ServerRelativeUrl").catch(function () { return { d: { Title: siteUrl, ServerRelativeUrl: '' } }; })
    ]).then(function (results) {
      var subsites = results[0].d.results;
      var lists = results[1].d.results;
      var web = results[2].d;

      clearEl(container);

      // Breadcrumb
      var crumb = create('span', { textContent: web.Title, onClick: function () {
        loadSiteBrowser(container, getSiteUrl());
      }});
      breadcrumb.appendChild(crumb);
      container.appendChild(breadcrumb);

      // Subsites
      subsites.forEach(function (sub) {
        var item = create('div', { className: 'browser-item', onClick: function () {
          var subUrl = window.location.origin + sub.ServerRelativeUrl;
          loadSiteBrowser(container, subUrl);
        }});
        item.appendChild(create('span', { className: 'icon', textContent: '\uD83D\uDCC1' }));
        item.appendChild(create('span', { textContent: sub.Title }));
        container.appendChild(item);
      });

      // Lists (filter out system lists)
      var customLists = lists.filter(function (l) {
        return l.BaseTemplate === 100 || l.BaseTemplate === 106 || l.BaseTemplate === 107;
      });

      customLists.forEach(function (lst) {
        var item = create('div', { className: 'browser-item', onClick: function () {
          // Select this list
          var siblings = container.querySelectorAll('.browser-item');
          siblings.forEach(function (s) { s.classList.remove('selected'); });
          item.classList.add('selected');
          setupData.listTitle = lst.Title;
          setupData.listId = lst.Id;
          setupData.siteUrl = siteUrl;
          setupData.listUrl = siteUrl + "/_api/web/lists(guid'" + lst.Id + "')";
          var inp = el('inputListUrl');
          if (inp) inp.value = siteUrl + '/Lists/' + encodeURIComponent(lst.Title);
        }});
        item.appendChild(create('span', { className: 'icon', textContent: '\uD83D\uDCCB' }));
        var titleSpan = create('span', { textContent: lst.Title });
        item.appendChild(titleSpan);
        item.appendChild(create('span', {
          textContent: '(' + lst.ItemCount + ' items)',
          style: 'color:var(--color-text-tertiary);font-size:11px;margin-left:auto;'
        }));
        container.appendChild(item);
      });

      if (subsites.length === 0 && customLists.length === 0) {
        container.appendChild(create('div', { className: 'browser-loading', textContent: 'No subsites or lists found at this location.' }));
      }
    }).catch(function (err) {
      clearEl(container);
      container.appendChild(create('div', { className: 'msg-error', textContent: 'Failed to load site contents. Check the URL and try again.' }));
    });
  }

  function validateStep1() {
    var errEl = el('step1Error');
    clearEl(errEl);

    var inputUrl = el('inputListUrl').value.trim();

    if (!setupData.listTitle && !inputUrl) {
      errEl.appendChild(create('div', { className: 'msg-error', textContent: 'Please enter a list URL or select a list from the browser.' }));
      return;
    }

    // If user typed a URL but didn't browse
    if (inputUrl && !setupData.listTitle) {
      // Try to parse the URL
      var parsed = parseListUrl(inputUrl);
      if (parsed) {
        setupData.siteUrl = parsed.siteUrl;
        setupData.listTitle = parsed.listTitle;
      } else {
        errEl.appendChild(create('div', { className: 'msg-error', textContent: 'Could not parse the list URL. Please use the browser to select a list.' }));
        return;
      }
    }

    // Validate by fetching fields
    errEl.appendChild(create('div', { className: 'msg-info', textContent: 'Validating list connection...' }));

    var escapedTitle = setupData.listTitle ? setupData.listTitle.replace(/'/g, "''") : '';
    var listApiUrl = setupData.siteUrl + "/_api/web/lists/getByTitle('" + encodeURIComponent(escapedTitle) + "')";
    if (setupData.listId) {
      listApiUrl = setupData.siteUrl + "/_api/web/lists(guid'" + setupData.listId + "')";
    }
    setupData.listUrl = listApiUrl;

    spGet(listApiUrl + "?$select=ListItemEntityTypeFullName,Id,Title")
      .then(function (data) {
        setupData.entityType = data.d.ListItemEntityTypeFullName;
        setupData.listId = data.d.Id;
        setupData.listTitle = data.d.Title;
        return loadFieldsAndSample();
      })
      .then(function () {
        setupStep = 1;
        renderSetupStep();
      })
      .catch(function (err) {
        clearEl(errEl);
        errEl.appendChild(create('div', { className: 'msg-error', textContent: 'Could not connect to the list. Please check the URL and your permissions.' }));
      });
  }

  function parseListUrl(url) {
    // Try to parse patterns like:
    // https://tenant.sharepoint.com/sites/SiteName/Lists/ListName
    // https://tenant.sharepoint.com/Lists/ListName
    var match = url.match(/(https?:\/\/.+?)\/Lists\/([^/?#]+)/i);
    if (match) {
      return { siteUrl: match[1], listTitle: decodeURIComponent(match[2]) };
    }
    // Try: .../sites/SiteName with a list title after
    match = url.match(/(https?:\/\/.+?\/sites\/[^/]+)/i);
    if (match) {
      // Extract list name from after /Lists/ if present
      var listMatch = url.match(/\/Lists\/([^/?#]+)/i);
      if (listMatch) {
        return { siteUrl: match[1], listTitle: decodeURIComponent(listMatch[1]) };
      }
    }
    return null;
  }

  function loadFieldsAndSample() {
    var listApiUrl = setupData.listUrl;
    return Promise.all([
      spGet(listApiUrl + "/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=Title,InternalName,TypeAsString,TypeDisplayName,FieldTypeKind"),
      spGet(listApiUrl + "/items?$top=1")
    ]).then(function (results) {
      setupData.fields = results[0].d.results.filter(function (f) {
        // Filter out system-ish fields
        var skip = ['ContentType','Attachments','Edit','DocIcon','ItemChildCount','FolderChildCount',
                     '_ComplianceFlags','_ComplianceTag','_ComplianceTagWrittenTime','_ComplianceTagUserId',
                     '_IsRecord','AppAuthor','AppEditor','ComplianceAssetId'];
        return skip.indexOf(f.InternalName) === -1;
      });
      setupData.sampleItem = results[1].d.results.length > 0 ? results[1].d.results[0] : null;
      setupData.hasDateFields = setupData.fields.some(function (f) { return f.TypeAsString === 'DateTime'; });
    });
  }

  // ----- Step 2: Map Fields -----
  function formatSampleValue(field, sampleItem) {
    if (!sampleItem || sampleItem[field.InternalName] == null) return '';
    var raw = sampleItem[field.InternalName];
    // Handle Choice/MultiChoice fields that come as objects
    if (typeof raw === 'object' && raw !== null) {
      if (raw.results && Array.isArray(raw.results)) return raw.results.join(', ');
      if (raw.__deferred) return '(lookup)';
      try { return JSON.stringify(raw); } catch (e) { return ''; }
    }
    return String(raw);
  }

  function guessCategory(field) {
    // Auto-detect start/end date fields by name matching
    if (field.TypeAsString !== 'DateTime') return 'none';
    var name = (field.Title + field.InternalName).replace(/[\s_\-]/g, '').toLowerCase();
    var startPatterns = ['startdate', 'datestart', 'calendarstart', 'calendarstartdate', 'begin', 'begindate', 'from', 'fromdate'];
    var endPatterns = ['enddate', 'dateend', 'calendarend', 'calendarenddate', 'duedate', 'due', 'finish', 'finishdate', 'to', 'todate', 'deadline'];
    for (var i = 0; i < startPatterns.length; i++) {
      if (name.indexOf(startPatterns[i]) !== -1) return 'startDate';
    }
    for (var i = 0; i < endPatterns.length; i++) {
      if (name.indexOf(endPatterns[i]) !== -1) return 'endDate';
    }
    return 'none';
  }

  function renderSetupStep2(body, footer) {
    el('setupTitle').textContent = 'Map Fields';
    el('setupSubtitle').textContent = 'Assign a category to each field you want to show in the calendar.';

    // Field table
    var wrapper = create('div', { style: 'max-height:350px;overflow-y:auto;overflow-x:hidden;border:1px solid var(--color-border);border-radius:var(--radius-md);' });
    var table = create('table', { className: 'field-table' });
    var thead = create('thead');
    var hrow = create('tr');
    ['Field Name', 'Type', 'Sample Value', 'Category', 'Color Base'].forEach(function (h) {
      var th = create('th', { textContent: h });
      if (h === 'Color Base') {
        th.style.textAlign = 'center';
      }
      hrow.appendChild(th);
    });
    thead.appendChild(hrow);
    table.appendChild(thead);

    var tbody = create('tbody', { id: 'fieldTableBody' });
    setupData.fields.forEach(function (field, idx) {
      var tr = create('tr');
      tr.appendChild(create('td', { textContent: field.Title }));
      tr.appendChild(create('td', { textContent: field.TypeDisplayName || field.TypeAsString }));

      var sampleVal = formatSampleValue(field, setupData.sampleItem);
      var sampleTd = create('td');
      sampleTd.appendChild(create('span', { className: 'sample-value', textContent: sampleVal, title: sampleVal }));
      tr.appendChild(sampleTd);

      // Category dropdown (merged Role + Display Style)
      var catTd = create('td');
      var catSelect = create('select', { 'data-field': field.InternalName, 'data-idx': String(idx), className: 'role-select' });
      catSelect.appendChild(create('option', { value: 'none', textContent: 'None' }));
      if (field.TypeAsString === 'DateTime') {
        catSelect.appendChild(create('option', { value: 'startDate', textContent: 'Start Date' }));
        catSelect.appendChild(create('option', { value: 'endDate', textContent: 'End Date' }));
      }
      catSelect.appendChild(create('option', { value: 'heading', textContent: 'Heading' }));
      catSelect.appendChild(create('option', { value: 'subheading', textContent: 'Sub-heading' }));
      catSelect.appendChild(create('option', { value: 'label', textContent: 'Label' }));
      catSelect.appendChild(create('option', { value: 'regular', textContent: 'Regular' }));

      // Determine initial value: revisit > auto-detect > none
      var existing = setupData.fieldMappings.find(function (m) { return m.internalName === field.InternalName; });
      if (existing) {
        catSelect.value = existing.style || 'regular';
      } else if (setupData.startDateField === field.InternalName) {
        catSelect.value = 'startDate';
      } else if (setupData.endDateField === field.InternalName) {
        catSelect.value = 'endDate';
      } else {
        // Auto-detect on first visit
        var guess = guessCategory(field);
        if (guess !== 'none') catSelect.value = guess;
      }

      catTd.appendChild(catSelect);
      tr.appendChild(catTd);

      // Color Base checkbox
      var cbTd = create('td', { style: 'text-align:center;' });
      var cbInput = create('input', {
        type: 'checkbox',
        className: 'color-base-checkbox',
        'data-field': field.InternalName,
        onChange: function (e) {
          if (e.target.checked) {
            document.querySelectorAll('.color-base-checkbox').forEach(function (other) {
              if (other !== e.target) other.checked = false;
            });
          }
        }
      });
      if (setupData.colorBaseField === field.InternalName) {
        cbInput.checked = true;
      } else if (!setupData.colorBaseField && APP.config && APP.config.colorBaseField === field.InternalName) {
        cbInput.checked = true;
      }
      cbTd.appendChild(cbInput);
      tr.appendChild(cbTd);

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrapper.appendChild(table);
    body.appendChild(wrapper);

    // Create date fields prompt
    if (!setupData.hasDateFields) {
      var prompt = create('div', { className: 'create-fields-prompt', id: 'createFieldsPrompt' });
      prompt.appendChild(create('h4', { textContent: 'No date fields found' }));
      prompt.appendChild(create('p', { textContent: 'This list has no DateTime fields. Would you like to create CalendarStartDate and CalendarEndDate fields?' }));
      var btnCreate = create('button', {
        className: 'btn btn-primary btn-sm',
        textContent: 'Create Date Fields',
        onClick: function () { createDateFields(prompt); }
      });
      prompt.appendChild(btnCreate);
      body.appendChild(prompt);
    }

    body.appendChild(create('div', { id: 'step2Error' }));

    // Footer
    var btnBack = create('button', {
      className: 'btn btn-ghost',
      textContent: '\u2190 Back',
      onClick: function () { setupStep = 0; renderSetupStep(); }
    });
    var btnNext = create('button', {
      className: 'btn btn-primary',
      textContent: 'Next \u2192',
      onClick: function () { validateStep2(); }
    });
    footer.appendChild(btnBack);
    footer.appendChild(btnNext);
  }

  function createDateFields(promptEl) {
    clearEl(promptEl);
    promptEl.appendChild(create('div', { className: 'msg-info', textContent: 'Creating date fields...' }));

    var listApiUrl = setupData.listUrl;

    Promise.all([
      spPost(listApiUrl + '/fields', {
        '__metadata': { 'type': 'SP.FieldDateTime' },
        'FieldTypeKind': 4,
        'Title': 'CalendarStartDate',
        'DisplayFormat': 1
      }),
      spPost(listApiUrl + '/fields', {
        '__metadata': { 'type': 'SP.FieldDateTime' },
        'FieldTypeKind': 4,
        'Title': 'CalendarEndDate',
        'DisplayFormat': 1
      })
    ]).then(function () {
      // Reload fields
      return loadFieldsAndSample();
    }).then(function () {
      setupData.hasDateFields = true;
      appendCSVRow('FIELD_CREATED', '', 'CalendarStartDate', 'DateTime', 'Created CalendarStartDate field');
      appendCSVRow('FIELD_CREATED', '', 'CalendarEndDate', 'DateTime', 'Created CalendarEndDate field');
      // Re-render step 2
      renderSetupStep();
    }).catch(function (err) {
      clearEl(promptEl);
      promptEl.appendChild(create('div', { className: 'msg-error', textContent: 'Failed to create fields. You may not have permission to modify this list.' }));
    });
  }

  function validateStep2() {
    var errEl = el('step2Error');
    clearEl(errEl);

    // Collect mappings from the merged Category dropdown
    var mappings = [];
    var startField = null;
    var endField = null;
    var catSelects = document.querySelectorAll('.role-select');
    var startCount = 0;
    var endCount = 0;

    catSelects.forEach(function (sel) {
      var fieldName = sel.getAttribute('data-field');
      var idx = parseInt(sel.getAttribute('data-idx'));
      var cat = sel.value;
      if (cat === 'none') return;
      if (cat === 'startDate') { startField = fieldName; startCount++; return; }
      if (cat === 'endDate') { endField = fieldName; endCount++; return; }
      // heading, subheading, label, regular are display styles
      var fieldObj = setupData.fields[idx];
      mappings.push({
        internalName: fieldName,
        title: fieldObj.Title,
        style: cat,
        type: fieldObj.TypeAsString
      });
    });

    if (startCount > 1 || endCount > 1) {
      errEl.appendChild(create('div', { className: 'msg-error', textContent: 'You can only select one Start Date and one End Date field.' }));
      return;
    }

    if (mappings.length === 0) {
      errEl.appendChild(create('div', { className: 'msg-error', textContent: 'Please select at least one field to display.' }));
      return;
    }

    setupData.fieldMappings = mappings;
    setupData.startDateField = startField;
    setupData.endDateField = endField;

    var colorBaseField = null;
    var cbCheckboxes = document.querySelectorAll('.color-base-checkbox');
    cbCheckboxes.forEach(function (cb) {
      if (cb.checked) {
        colorBaseField = cb.getAttribute('data-field');
      }
    });

    setupData.colorBaseField = colorBaseField;
    if (colorBaseField) {
      clearEl(errEl);
      var spinner = create('div', { className: 'msg-info', style: 'display:flex;align-items:center;gap:8px;' }, [
        create('span', { className: 'spinner', style: 'width:14px;height:14px;border-width:2px;' }),
        'Analyzing field values and assigning colors...'
      ]);
      errEl.appendChild(spinner);

      var listApiUrl = setupData.listUrl;
      var expandParam = '';
      var fieldObj = setupData.fields.find(function (f) { return f.InternalName === colorBaseField; });
      if (fieldObj && (fieldObj.TypeAsString === 'User' || fieldObj.TypeAsString === 'UserMulti' || fieldObj.TypeAsString === 'Lookup' || fieldObj.TypeAsString === 'LookupMulti')) {
        expandParam = '&$select=*,' + colorBaseField + '/Title&$expand=' + colorBaseField;
      }

      spGet(listApiUrl + "/items?$top=500" + expandParam)
        .then(function (data) {
          var items = data && data.d && data.d.results ? data.d.results : [];
          var uniqueVals = getUniqueValuesForField(items, colorBaseField);
          
          var newColorMap = {};
          var existingMap = setupData.colorMap || (APP.config && APP.config.colorMap) || {};
          var existingBase = APP.config && APP.config.colorBaseField;

          uniqueVals.forEach(function (val, idx) {
            if (colorBaseField === existingBase && existingMap[val]) {
              newColorMap[val] = existingMap[val];
            } else {
              newColorMap[val] = PRESET_COLORS[idx % PRESET_COLORS.length];
            }
          });

          setupData.uniqueValues = uniqueVals;
          setupData.colorMap = newColorMap;

          setupStep = 2;
          renderSetupStep();
        })
        .catch(function (err) {
          clearEl(errEl);
          errEl.appendChild(create('div', { className: 'msg-error', textContent: 'Failed to retrieve list items for color mapping. Please try again.' }));
        });
    } else {
      setupData.colorMap = null;
      setupData.uniqueValues = [];
      setupStep = 3;
      renderSetupStep();
    }
  }

  // ----- Step 2.5 (New Step 2): Color Configuration -----
  function renderSetupStep2ColorConfig(body, footer) {
    el('setupTitle').textContent = 'Configure Colors';
    el('setupSubtitle').textContent = 'Customize colors for each unique value in the "' + getFieldTitle(setupData.colorBaseField) + '" field.';

    var wrapper = create('div', { 
      style: 'max-height:300px;overflow-y:auto;overflow-x:hidden;border:1px solid var(--color-border);border-radius:var(--radius-md);padding:16px;background:var(--color-bg);' 
    });

    if (!setupData.uniqueValues || setupData.uniqueValues.length === 0) {
      wrapper.appendChild(create('div', { className: 'msg-info', textContent: 'No unique values found in the selected field. You can proceed, and colors will be assigned dynamically as items are added.' }));
    } else {
      var grid = create('div', { 
        style: 'display:grid;grid-template-columns:1fr auto;gap:12px;align-items:center;' 
      });

      setupData.uniqueValues.forEach(function (val) {
        var label = create('span', { 
          textContent: val, 
          style: 'font-weight:500;font-size:13px;color:var(--color-text);text-overflow:ellipsis;overflow:hidden;white-space:nowrap;' 
        });
        
        var picker = create('input', {
          type: 'color',
          className: 'color-picker-input',
          value: setupData.colorMap[val] || '#0078d4',
          style: 'width:48px;height:32px;border:1px solid var(--color-border);border-radius:var(--radius-sm);cursor:pointer;padding:2px;background:transparent;',
          onChange: function (e) {
            setupData.colorMap[val] = e.target.value;
          }
        });

        grid.appendChild(label);
        grid.appendChild(picker);
      });

      wrapper.appendChild(grid);
    }

    body.appendChild(wrapper);

    // Footer buttons
    var btnBack = create('button', {
      className: 'btn btn-ghost',
      textContent: '\u2190 Back',
      onClick: function () { setupStep = 1; renderSetupStep(); }
    });
    var btnNext = create('button', {
      className: 'btn btn-primary',
      textContent: 'Next \u2192',
      onClick: function () { setupStep = 3; renderSetupStep(); }
    });

    footer.appendChild(btnBack);
    footer.appendChild(btnNext);
  }

  // ----- Step 3: Page Title & Confirm -----
  function renderSetupStep3(body, footer) {
    el('setupTitle').textContent = 'Finalize Setup';
    el('setupSubtitle').textContent = 'Give your calendar a name and confirm settings.';

    // Page title
    var grp = create('div', { className: 'form-group' });
    grp.appendChild(create('label', { textContent: 'Calendar Page Title', for: 'inputPageTitle' }));
    var inp = create('input', {
      className: 'form-input',
      id: 'inputPageTitle',
      type: 'text',
      placeholder: 'e.g. Project Timeline',
      value: setupData.pageTitle || (setupData.listTitle ? setupData.listTitle + ' Calendar' : '')
    });
    grp.appendChild(inp);
    body.appendChild(grp);

    // Item Cap Limit
    var grpLimit = create('div', { className: 'form-group' });
    grpLimit.appendChild(create('label', { textContent: 'Maximum Items to Fetch', for: 'inputItemLimit' }));
    var inpLimit = create('input', {
      className: 'form-input',
      id: 'inputItemLimit',
      type: 'number',
      min: '100',
      max: '50000',
      step: '100',
      placeholder: 'e.g. 5000',
      value: setupData.itemLimit || 5000
    });
    grpLimit.appendChild(inpLimit);
    body.appendChild(grpLimit);

    // Summary
    var summary = create('div', {
      style: 'background:var(--color-bg-secondary);border:1px solid var(--color-border);border-radius:var(--radius-md);padding:16px;'
    });
    summary.appendChild(create('h4', { textContent: 'Configuration Summary', style: 'font-size:14px;font-weight:600;margin-bottom:12px;' }));

    var items = [
      ['List', setupData.listTitle],
      ['Start Date Field', setupData.startDateField || '(none)'],
      ['End Date Field', setupData.endDateField || '(none)'],
      ['Display Fields', setupData.fieldMappings.map(function (m) { return m.title + ' (' + m.style + ')'; }).join(', ')]
    ];
    items.forEach(function (item) {
      var row = create('div', { style: 'display:flex;gap:8px;margin-bottom:6px;font-size:13px;' });
      row.appendChild(create('span', { textContent: item[0] + ':', style: 'font-weight:600;min-width:120px;color:var(--color-text-secondary);' }));
      row.appendChild(create('span', { textContent: item[1] }));
      summary.appendChild(row);
    });

    if (!setupData.startDateField || !setupData.endDateField) {
      summary.appendChild(create('div', { className: 'msg-info', textContent: 'Note: Without start/end date fields, the calendar will be empty. Items will appear only in the card panel.', style: 'margin-top:12px;' }));
    }

    body.appendChild(summary);

    // Footer
    var btnBack = create('button', {
      className: 'btn btn-ghost',
      textContent: '\u2190 Back',
      onClick: function () {
        if (setupData.colorBaseField) {
          setupStep = 2;
        } else {
          setupStep = 1;
        }
        renderSetupStep();
      }
    });
    var btnSave = create('button', {
      className: 'btn btn-primary',
      textContent: 'Save & Launch',
      onClick: function () { finalizeSetup(); }
    });
    footer.appendChild(btnBack);
    footer.appendChild(btnSave);
  }

  function finalizeSetup() {
    var title = el('inputPageTitle').value.trim() || 'SharePoint Calendar';
    var limit = parseInt(el('inputItemLimit').value, 10);
    if (isNaN(limit) || limit < 100) limit = 5000;
    else if (limit > 50000) limit = 50000;
    setupData.pageTitle = title;
    setupData.itemLimit = limit;

    var colorBaseFieldType = null;
    if (setupData.colorBaseField) {
      var fieldObj = setupData.fields.find(function (f) { return f.InternalName === setupData.colorBaseField; });
      if (fieldObj) colorBaseFieldType = fieldObj.TypeAsString;
    }

    // Build config
    APP.config = {
      listUrl: setupData.listUrl,
      siteUrl: setupData.siteUrl,
      listTitle: setupData.listTitle,
      listId: setupData.listId,
      fieldMappings: setupData.fieldMappings,
      startDateField: setupData.startDateField,
      endDateField: setupData.endDateField,
      pageTitle: title,
      entityType: setupData.entityType,
      itemLimit: limit,
      colorBaseField: setupData.colorBaseField || null,
      colorBaseFieldType: colorBaseFieldType || null,
      colorMap: setupData.colorMap || null
    };

    // Prune existing SETUP_CONFIG rows to keep log file size small and compact
    if (Array.isArray(APP.csvRows)) {
      APP.csvRows = APP.csvRows.filter(function (row) {
        return row.ActionType !== 'SETUP_CONFIG';
      });
    } else {
      APP.csvRows = [];
    }

    // Save to CSV
    appendCSVRow('SETUP_CONFIG', '', 'listUrl', JSON.stringify(APP.config.listUrl), '');
    appendCSVRow('SETUP_CONFIG', '', 'siteUrl', JSON.stringify(APP.config.siteUrl), '');
    appendCSVRow('SETUP_CONFIG', '', 'listTitle', JSON.stringify(APP.config.listTitle), '');
    appendCSVRow('SETUP_CONFIG', '', 'listId', JSON.stringify(APP.config.listId), '');
    appendCSVRow('SETUP_CONFIG', '', 'fieldMappings', JSON.stringify(APP.config.fieldMappings), '');
    appendCSVRow('SETUP_CONFIG', '', 'startDateField', JSON.stringify(APP.config.startDateField), '');
    appendCSVRow('SETUP_CONFIG', '', 'endDateField', JSON.stringify(APP.config.endDateField), '');
    appendCSVRow('SETUP_CONFIG', '', 'pageTitle', JSON.stringify(APP.config.pageTitle), '');
    appendCSVRow('SETUP_CONFIG', '', 'entityType', JSON.stringify(APP.config.entityType), '');
    appendCSVRow('SETUP_CONFIG', '', 'itemLimit', JSON.stringify(APP.config.itemLimit), '');
    appendCSVRow('SETUP_CONFIG', '', 'colorBaseField', JSON.stringify(APP.config.colorBaseField), '');
    appendCSVRow('SETUP_CONFIG', '', 'colorBaseFieldType', JSON.stringify(APP.config.colorBaseFieldType), '');
    appendCSVRow('SETUP_CONFIG', '', 'colorMap', JSON.stringify(APP.config.colorMap), '');

    showSaving();
    saveCSV().then(function () {
      showSaved();
      // Remove setup overlay
      var overlay = el('setupOverlay');
      if (overlay) overlay.remove();
      launchApp();
    }).catch(function () {
      showSaveError();
    });
  }

  // ===== Main Application =====
  // ===== Main Application =====
  function setView(view) {
    APP.currentView = view;
    localStorage.setItem('SPCalendar_PreferredView', view);

    var btnMonth = el('btnViewMonthGrid');
    var btnTimeline = el('btnViewTimeline');
    if (btnMonth && btnTimeline) {
      if (view === 'month') {
        btnMonth.className = 'btn btn-primary btn-sm';
        btnTimeline.className = 'btn btn-ghost btn-sm';
      } else {
        btnMonth.className = 'btn btn-ghost btn-sm';
        btnTimeline.className = 'btn btn-primary btn-sm';
      }
    }

    var start = APP.calendarStartDate || new Date();
    if (view === 'month') {
      APP.calendarStartDate = new Date(start.getFullYear(), start.getMonth(), 1);
      APP.calendarEndDate = new Date(start.getFullYear(), start.getMonth() + 1, 0);
    } else {
      APP.calendarStartDate = new Date(start.getFullYear(), start.getMonth(), 1);
      APP.calendarEndDate = new Date(start.getFullYear(), start.getMonth() + 2, 0);
    }

    renderCalendar();
    renderCards();

    if (view === 'timeline') {
      scrollToToday();
    }
  }

  function launchApp() {
    el('pageTitle').textContent = APP.config.pageTitle;
    document.title = APP.config.pageTitle;
    show(el('mainContainer'));
    show(el('btnReconfigure'));
    show(el('btnEmbed'));

    if (APP.isEmbedMode) {
      el('pageHeader').classList.add('embed-mode');
    }

    // Set initial calendar range: complete month(s) based on current view
    var now = new Date();
    if (APP.currentView === 'month') {
      APP.calendarStartDate = new Date(now.getFullYear(), now.getMonth(), 1);
      APP.calendarEndDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    } else {
      APP.calendarStartDate = new Date(now.getFullYear(), now.getMonth(), 1);
      APP.calendarEndDate = new Date(now.getFullYear(), now.getMonth() + 2, 0); // end of next month
    }

    // Update active button state
    var btnMonth = el('btnViewMonthGrid');
    var btnTimeline = el('btnViewTimeline');
    if (btnMonth && btnTimeline) {
      if (APP.currentView === 'month') {
        btnMonth.className = 'btn btn-primary btn-sm';
        btnTimeline.className = 'btn btn-ghost btn-sm';
      } else {
        btnMonth.className = 'btn btn-ghost btn-sm';
        btnTimeline.className = 'btn btn-primary btn-sm';
      }
    }

    loadItems().then(function () {
      APP._lastItemsFingerprint = itemsFingerprint();
      renderCalendar();
      renderCards();
      startPolling();
      if (APP.currentView === 'timeline') {
        scrollToToday();
      }
    });
  }

  function scrollToToday() {
    if (APP.currentView !== 'timeline') return;
    // Auto-scroll so today is visible with 3 days before it
    var now = new Date();
    var dayOffset = daysBetween(APP.calendarStartDate, now);
    var scrollPos = Math.max(0, (dayOffset - 3) * CELL_WIDTH);
    var wrapper = el('calendarScrollWrapper');
    if (wrapper) wrapper.scrollLeft = scrollPos;
  }

  // ===== Load Items =====
  function buildSelectFields() {
    var fields = ['Id'];
    if (APP.config.startDateField) fields.push(APP.config.startDateField);
    if (APP.config.endDateField) fields.push(APP.config.endDateField);
    APP.config.fieldMappings.forEach(function (m) {
      if (fields.indexOf(m.internalName) === -1) fields.push(m.internalName);
    });
    return fields.join(',');
  }

  function loadItems() {
    var baseUrl = APP.config.listUrl + "/items";

    // Build $expand and $select for User/Lookup fields
    var expandFields = [];
    APP.config.fieldMappings.forEach(function (m) {
      if (m.type === 'User' || m.type === 'UserMulti' || m.type === 'Lookup' || m.type === 'LookupMulti') {
        if (expandFields.indexOf(m.internalName) === -1) expandFields.push(m.internalName);
      }
    });
    if (APP.config.colorBaseField && APP.config.colorBaseFieldType) {
      var t = APP.config.colorBaseFieldType;
      if (t === 'User' || t === 'UserMulti' || t === 'Lookup' || t === 'LookupMulti') {
        if (expandFields.indexOf(APP.config.colorBaseField) === -1) {
          expandFields.push(APP.config.colorBaseField);
        }
      }
    }
    var expandSelectParam = '';
    if (expandFields.length > 0) {
      var selectParts = ['*'];
      expandFields.forEach(function (f) {
        selectParts.push(f + '/Title');
        selectParts.push(f + '/EMail');
      });
      expandSelectParam = '&$select=' + selectParts.join(',') + '&$expand=' + expandFields.join(',');
    }

    var limit = (APP.config && APP.config.itemLimit) ? APP.config.itemLimit : 5000;
    var pageSize = Math.min(100, limit);
    var url = baseUrl + "?$top=" + pageSize + "&$orderby=Modified desc" + expandSelectParam;

    return spGet(url).then(function (data) {
      var items = data && data.d && data.d.results ? data.d.results : [];
      if (items.length > limit) {
        items = items.slice(0, limit);
      }
      console.log('[SPCalendar] Loaded', items.length, 'items');
      APP.items = items;
      if (data && data.d && data.d.__next && APP.items.length < limit) {
        return loadAllPages(data.d.__next);
      }
    }).catch(function (err) {
      console.error('[SPCalendar] Error loading items:', err);
      APP.items = [];
    });
  }

  function loadAllPages(nextUrl) {
    var limit = (APP.config && APP.config.itemLimit) ? APP.config.itemLimit : 5000;
    return spGet(nextUrl).then(function (data) {
      var items = data && data.d && data.d.results ? data.d.results : [];
      var combined = APP.items.concat(items);
      if (combined.length > limit) {
        combined = combined.slice(0, limit);
      }
      APP.items = combined;
      if (data && data.d && data.d.__next && APP.items.length < limit) {
        return loadAllPages(data.d.__next);
      }
    }).catch(function (err) {
      console.error('[SPCalendar] Error loading next pages:', err);
    });
  }

  // Helper to clear dynamically rendered elements from calendarBody without wiping grid lines
  function clearCalendarBodyDynamicElements() {
    var calBody = el('calendarBody');
    if (!calBody) return;
    calBody.querySelectorAll('.week-row').forEach(function (r) { r.remove(); });
    calBody.querySelectorAll('.calendar-block').forEach(function (b) { b.remove(); });
  }

  // ===== Calendar Rendering =====
  // ===== Calendar Rendering =====
  function renderCalendar() {
    var start = APP.calendarStartDate;
    var end = APP.calendarEndDate;
    var today = new Date();

    // Toggle view classes on the scroll wrapper
    var scrollWrapper = el('calendarScrollWrapper');
    if (scrollWrapper) {
      if (APP.currentView === 'month') {
        scrollWrapper.classList.remove('view-timeline');
        scrollWrapper.classList.add('view-month');
      } else {
        scrollWrapper.classList.remove('view-month');
        scrollWrapper.classList.add('view-timeline');
      }
    }

    // Update range label: show just Month Year if spanning exactly 1 month
    if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
      el('dateRangeLabel').textContent = MONTH_FULL[start.getMonth()] + ' ' + start.getFullYear();
    } else {
      el('dateRangeLabel').textContent = MONTH_FULL[start.getMonth()] + ' ' + start.getFullYear() +
        ' - ' + MONTH_FULL[end.getMonth()] + ' ' + end.getFullYear();
    }

    if (APP.currentView === 'month') {
      renderMonthViewGrid(start, end, today);
    } else {
      clearCalendarBodyDynamicElements();
      var totalDays = daysBetween(start, end) + 1;

      // Month header
      var monthHeader = el('calendarMonthHeader');
      show(monthHeader);
      clearEl(monthHeader);
      var curMonth = -1;
      var monthStartIdx = 0;
      var months = [];
      for (var d = 0; d < totalDays; d++) {
        var date = addDays(start, d);
        var m = date.getMonth();
        if (m !== curMonth) {
          if (curMonth !== -1) {
            months.push({ label: MONTH_NAMES[curMonth] + ' ' + addDays(start, monthStartIdx).getFullYear(), span: d - monthStartIdx });
          }
          curMonth = m;
          monthStartIdx = d;
        }
      }
      months.push({ label: MONTH_NAMES[curMonth] + ' ' + addDays(start, monthStartIdx).getFullYear(), span: totalDays - monthStartIdx });

      months.forEach(function (mo) {
        var cell = create('div', { className: 'calendar-month-cell', textContent: mo.label });
        cell.style.width = (mo.span * CELL_WIDTH) + 'px';
        cell.style.minWidth = cell.style.width;
        monthHeader.appendChild(cell);
      });

      // Day header
      var dayHeader = el('calendarDayHeader');
      show(dayHeader);
      clearEl(dayHeader);
      for (var d = 0; d < totalDays; d++) {
        var date = addDays(start, d);
        var cls = 'calendar-day-cell';
        if (isWeekend(date)) cls += ' weekend';
        if (isSameDay(date, today)) cls += ' today';
        var cell = create('div', { className: cls, textContent: date.getDate() });
        dayHeader.appendChild(cell);
      }

      // Grid lines
      var gridLines = el('calendarGridLines');
      show(gridLines);
      clearEl(gridLines);
      for (var d = 0; d < totalDays; d++) {
        var date = addDays(start, d);
        var cls = 'calendar-grid-line';
        if (isWeekend(date)) cls += ' weekend';
        if (isSameDay(date, today)) cls += ' today';
        gridLines.appendChild(create('div', { className: cls }));
      }

      // Render blocks
      renderBlocks(start, end, totalDays);
    }
  }

  function renderMonthViewGrid(start, end, today) {
    var body = el('calendarBody');
    clearCalendarBodyDynamicElements();
    
    // Hide grid lines in DOM
    var gridLines = el('calendarGridLines');
    if (gridLines) hide(gridLines);
    
    var monthHeader = el('calendarMonthHeader');
    if (monthHeader) hide(monthHeader);

    // Render weekday headers starting from locale's first day of week
    var firstDayOfWeek = getLocaleFirstDayOfWeek();
    var weekdayNames = getLocaleWeekdayNames();
    
    var dayHeader = el('calendarDayHeader');
    show(dayHeader);
    clearEl(dayHeader);
    for (var i = 0; i < 7; i++) {
      var dayIndex = (firstDayOfWeek + i) % 7;
      var name = weekdayNames[dayIndex];
      var isWk = dayIndex === 0 || dayIndex === 6;
      var cls = 'calendar-day-name-cell';
      if (isWk) cls += ' weekend';
      var cell = create('div', { className: cls, textContent: name });
      dayHeader.appendChild(cell);
    }
    
    // Determine month boundaries and padding days
    var firstOfMonth = new Date(start.getFullYear(), start.getMonth(), 1);
    var firstDayIndex = firstOfMonth.getDay();
    var padDays = (firstDayIndex - firstDayOfWeek + 7) % 7;
    var gridStart = addDays(firstOfMonth, -padDays);
    
    var lastOfMonth = new Date(end.getFullYear(), end.getMonth() + 1, 0);
    var totalDaysCovered = daysBetween(gridStart, lastOfMonth) + 1;
    var totalGridDays = Math.ceil(totalDaysCovered / 7) * 7;
    
    // Render week rows
    var totalWeeks = totalGridDays / 7;
    for (var w = 0; w < totalWeeks; w++) {
      var weekStart = addDays(gridStart, w * 7);
      var weekEnd = addDays(weekStart, 6);
      
      var weekRow = create('div', { className: 'week-row' });
      
      // Render 7 background cells
      var weekGridDays = create('div', { className: 'week-grid-days' });
      for (var d = 0; d < 7; d++) {
        var dayDate = addDays(weekStart, d);
        var cls = 'week-day-cell';
        if (isWeekend(dayDate)) cls += ' weekend';
        if (isSameDay(dayDate, today)) cls += ' today';
        if (dayDate.getMonth() !== start.getMonth()) cls += ' outside-month';
        
        var cell = create('div', { className: cls });
        
        // Day number label
        var numContainer = create('div', { className: 'day-number-container' });
        var numLabel = create('span', { className: 'day-number-label', textContent: dayDate.getDate() });
        numContainer.appendChild(numLabel);
        cell.appendChild(numContainer);
        
        weekGridDays.appendChild(cell);
      }
      weekRow.appendChild(weekGridDays);
      
      // Create week events container
      var evContainer = create('div', { className: 'week-events-container' });
      weekRow.appendChild(evContainer);
      
      body.appendChild(weekRow);
      
      // Render event blocks for this week row
      renderMonthGridBlocksForWeek(evContainer, weekStart, weekEnd);
    }
    
    // Stack all blocks across the entire month grid after layout pass
    setTimeout(restackMonthGridBlocks, 0);
  }

  function renderMonthGridBlocksForWeek(evContainer, weekStart, weekEnd) {
    if (!APP.config.startDateField || !APP.config.endDateField) return;

    var itemsInWeek = APP.items.filter(function (item) {
      var s = parseDate(item[APP.config.startDateField]);
      var e = parseDate(item[APP.config.endDateField]);
      return s && e && s <= weekEnd && e >= weekStart;
    }).map(function (item) {
      return {
        id: item.Id,
        start: parseDate(item[APP.config.startDateField]),
        end: parseDate(item[APP.config.endDateField]),
        item: item
      };
    });

    // Sort by start date, then duration, then ID
    itemsInWeek.sort(function (a, b) {
      var startDiff = a.start - b.start;
      if (startDiff !== 0) return startDiff;
      var durationA = a.end - a.start;
      var durationB = b.end - b.start;
      return durationB - durationA; // longer duration first
    });

    itemsInWeek.forEach(function (entry) {
      var blockStart = entry.start < weekStart ? weekStart : entry.start;
      var blockEnd = entry.end > weekEnd ? weekEnd : entry.end;
      
      var startOffset = daysBetween(weekStart, blockStart);
      var spanDays = daysBetween(blockStart, blockEnd) + 1;
      
      var left = (startOffset * 14.285714);
      var width = (spanDays * 14.285714);
      
      var colorClass = 'block-color-' + (entry.id % 8);
      var customBg = null;
      if (APP.config && APP.config.colorBaseField && APP.config.colorMap) {
        var rawVal = entry.item[APP.config.colorBaseField];
        var strVal = formatFieldValue(rawVal);
        if (strVal && strVal.trim()) {
          var trimmedVal = strVal.trim();
          customBg = APP.config.colorMap[trimmedVal];
          if (!customBg) {
            var numKeys = Object.keys(APP.config.colorMap).length;
            customBg = PRESET_COLORS[numKeys % PRESET_COLORS.length];
            APP.config.colorMap[trimmedVal] = customBg;
          }
        }
      }

      var block = create('div', {
        className: 'calendar-block ' + (customBg ? '' : colorClass),
        'data-item-id': String(entry.id),
        'data-start-day': String(startOffset),
        'data-end-day': String(startOffset + spanDays - 1),
        onClick: function (e) {
          e.stopPropagation();
          selectItem(entry.id);
        }
      });
      block.style.left = left + '%';
      if (customBg) {
        block.style.backgroundColor = customBg;
        block.style.color = '#ffffff';
      }
      block.style.width = 'calc(' + width + '% - 4px)';
      
      if (APP.selectedItemId === entry.id || APP.addingDatesForId === entry.id) {
        block.classList.add('selected');
        block.classList.add('expanded');
      }
      
      // Render fields
      var labelsContainer = null;
      sortedMappings().forEach(function (mapping) {
        var val = entry.item[mapping.internalName];
        var valStr = formatFieldValue(val);
        if (!valStr) return;
        var span;
        if (mapping.style === 'heading') {
          span = create('span', { className: 'block-heading', textContent: valStr });
          block.appendChild(span);
        } else if (mapping.style === 'subheading') {
          span = create('span', { className: 'block-subheading', textContent: valStr });
          block.appendChild(span);
        } else if (mapping.style === 'label') {
          if (!labelsContainer) {
            labelsContainer = create('div', { className: 'block-labels-row' });
            block.appendChild(labelsContainer);
          }
          span = create('span', { className: 'block-label', textContent: valStr });
          labelsContainer.appendChild(span);
        } else {
          span = create('span', { className: 'block-regular', textContent: valStr });
          block.appendChild(span);
        }
      });
      
      evContainer.appendChild(block);
    });
  }

  function restackMonthGridBlocks() {
    var weekRows = document.querySelectorAll('.week-row');
    weekRows.forEach(function (row) {
      var blocks = Array.prototype.slice.call(row.querySelectorAll('.calendar-block'));
      if (blocks.length === 0) {
        var evCont = row.querySelector('.week-events-container');
        if (evCont) evCont.style.minHeight = '50px';
        return;
      }
      
      blocks.sort(function (a, b) {
        var startA = parseInt(a.getAttribute('data-start-day'), 10) || 0;
        var startB = parseInt(b.getAttribute('data-start-day'), 10) || 0;
        if (startA !== startB) return startA - startB;
        var endA = parseInt(a.getAttribute('data-end-day'), 10) || 0;
        var endB = parseInt(b.getAttribute('data-end-day'), 10) || 0;
        return endB - endA; // longer span first
      });
      
      var placed = [];
      var maxBottom = 0;
      
      blocks.forEach(function (block) {
        var startDay = parseInt(block.getAttribute('data-start-day'), 10) || 0;
        var endDay = parseInt(block.getAttribute('data-end-day'), 10) || 0;
        var measuredHeight = block.offsetHeight || 90;
        
        var topPos = 4;
        var collision = true;
        
        while (collision) {
          collision = false;
          for (var i = 0; i < placed.length; i++) {
            var p = placed[i];
            var horizOverlap = (startDay <= p.endDay && endDay >= p.startDay);
            if (horizOverlap) {
              var pBottom = p.topPos + p.measuredHeight + 4;
              if (topPos < pBottom && topPos + measuredHeight + 4 > p.topPos) {
                topPos = pBottom;
                collision = true;
                break;
              }
            }
          }
        }
        
        block.style.top = topPos + 'px';
        
        placed.push({
          startDay: startDay,
          endDay: endDay,
          topPos: topPos,
          measuredHeight: measuredHeight
        });
        
        var bottom = topPos + measuredHeight;
        if (bottom > maxBottom) maxBottom = bottom;
      });
      
      var evCont = row.querySelector('.week-events-container');
      if (evCont) {
        evCont.style.minHeight = (maxBottom + 10) + 'px';
      }
    });
  }

  function renderBlocks(calStart, calEnd, totalDays) {
    var calBody = el('calendarBody');
    // Remove old blocks
    calBody.querySelectorAll('.calendar-block').forEach(function (b) { b.remove(); });

    if (!APP.config.startDateField || !APP.config.endDateField) {
      calBody.style.minHeight = '200px';
      return;
    }

    // Filter items that have both dates and overlap with calendar range
    var itemsWithDates = APP.items.filter(function (item) {
      var s = parseDate(item[APP.config.startDateField]);
      var e = parseDate(item[APP.config.endDateField]);
      return s && e && s <= calEnd && e >= calStart;
    }).map(function (item) {
      return {
        id: item.Id,
        start: parseDate(item[APP.config.startDateField]),
        end: parseDate(item[APP.config.endDateField]),
        item: item
      };
    });

    // Sort by start date
    itemsWithDates.sort(function (a, b) { return a.start - b.start; });

    // --- Two-pass rendering: create blocks, measure, then stack ---

    // Pass 1: Create blocks and append to DOM (auto height for measurement)
    var blockEntries = [];
    itemsWithDates.forEach(function (entry) {
      var blockStart = entry.start < calStart ? calStart : entry.start;
      var blockEnd = entry.end > calEnd ? calEnd : entry.end;

      var startOffset = daysBetween(calStart, blockStart);
      var spanDays = daysBetween(blockStart, blockEnd) + 1;

      var left = startOffset * CELL_WIDTH;
      var width = spanDays * CELL_WIDTH - 2;

      var colorClass = 'block-color-' + (entry.id % 8);
      var customBg = null;
      if (APP.config && APP.config.colorBaseField && APP.config.colorMap) {
        var rawVal = entry.item[APP.config.colorBaseField];
        var strVal = formatFieldValue(rawVal);
        if (strVal && strVal.trim()) {
          var trimmedVal = strVal.trim();
          customBg = APP.config.colorMap[trimmedVal];
          if (!customBg) {
            var numKeys = Object.keys(APP.config.colorMap).length;
            customBg = PRESET_COLORS[numKeys % PRESET_COLORS.length];
            APP.config.colorMap[trimmedVal] = customBg;
          }
        }
      }

      var block = create('div', {
        className: 'calendar-block ' + (customBg ? '' : colorClass),
        'data-item-id': String(entry.id),
        onClick: function () { selectItem(entry.id); }
      });
      block.style.left = left + 'px';
      if (customBg) {
        block.style.backgroundColor = customBg;
        block.style.color = '#ffffff';
      }
      block.style.width = width + 'px';

      if (APP.selectedItemId === entry.id || APP.addingDatesForId === entry.id) {
        block.classList.add('selected');
        block.classList.add('expanded');
      }

      // Render fields: Heading > Sub-heading > Regular > Label
      var labelsContainer = null;
      sortedMappings().forEach(function (mapping) {
        var val = entry.item[mapping.internalName];
        var valStr = formatFieldValue(val);
        if (!valStr) return;
        var span;
        if (mapping.style === 'heading') {
          span = create('span', { className: 'block-heading', textContent: valStr });
          block.appendChild(span);
        } else if (mapping.style === 'subheading') {
          span = create('span', { className: 'block-subheading', textContent: valStr });
          block.appendChild(span);
        } else if (mapping.style === 'label') {
          if (!labelsContainer) {
            labelsContainer = create('div', { className: 'block-labels-row' });
            block.appendChild(labelsContainer);
          }
          span = create('span', { className: 'block-label', textContent: valStr });
          labelsContainer.appendChild(span);
        } else {
          span = create('span', { className: 'block-regular', textContent: valStr });
          block.appendChild(span);
        }
      });

      calBody.appendChild(block);
      blockEntries.push({ entry: entry, block: block });
    });

    // Store for restacking on expand/collapse
    APP._blockEntries = blockEntries;

    // Pass 2: Measure and stack
    restackBlocks();
  }

  function restackBlocks() {
    if (APP.currentView === 'month') {
      restackMonthGridBlocks();
      return;
    }

    var blockEntries = APP._blockEntries;
    if (!blockEntries || blockEntries.length === 0) return;

    var placedBlocks = [];

    blockEntries.forEach(function (be) {
      var entry = be.entry;
      var block = be.block;
      var measuredHeight = block.offsetHeight || 40;

      // Filter previously placed blocks that horizontally overlap with this entry
      var overlappingPlaced = placedBlocks.filter(function (placed) {
        return placed.start <= entry.end && placed.end >= entry.start;
      });

      // Find the highest vertical position where this block can fit
      var topPos = BLOCK_TOP_OFFSET;
      var collision = true;
      while (collision) {
        collision = false;
        for (var i = 0; i < overlappingPlaced.length; i++) {
          var A = overlappingPlaced[i];
          var aBottom = A.topPos + A.measuredHeight + BLOCK_GAP;
          // Check for vertical overlap between [topPos, topPos + measuredHeight] and [A.topPos, aBottom]
          if (topPos < aBottom && topPos + measuredHeight + BLOCK_GAP > A.topPos) {
            topPos = aBottom;
            collision = true;
            break; // Shifted vertically, re-run checks
          }
        }
      }

      entry.topPos = topPos;
      entry.measuredHeight = measuredHeight;
      block.style.top = topPos + 'px';

      placedBlocks.push({
        start: entry.start,
        end: entry.end,
        topPos: topPos,
        measuredHeight: measuredHeight
      });
    });

    var calBody = el('calendarBody');
    var maxBottom = 0;
    placedBlocks.forEach(function (pb) {
      var bottom = pb.topPos + pb.measuredHeight;
      if (bottom > maxBottom) maxBottom = bottom;
    });
    calBody.style.minHeight = Math.max(200, maxBottom + 20) + 'px';
  }

  function checkLimitBanner() {
    var banner = el('limitBanner');
    if (!banner || !APP.config) return;

    // Check if dismissed within the last 2 weeks
    var dismissedUntil = localStorage.getItem('SPCalendar_LimitDismissedUntil');
    if (dismissedUntil && Date.now() < parseInt(dismissedUntil, 10)) {
      hide(banner);
      return;
    }

    var limit = APP.config.itemLimit || 5000;
    if (APP.items && APP.items.length >= limit) {
      show(banner);
    } else {
      hide(banner);
    }
  }

  // ===== Card Panel =====
  function renderCards() {
    var container = el('sidebarScroll');
    clearEl(container);

    checkLimitBanner();

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var archivedIds = APP._archivedIds || {};
    var activeTab = APP.currentTab || 'scheduled';

    var scheduledItems = [];
    var pastDueItems = [];
    var unscheduledItems = [];
    var archivedItems = [];

    APP.items.forEach(function (item) {
      var isArchived = !!archivedIds[item.Id];
      if (isArchived) {
        archivedItems.push(item);
        return;
      }

      var startVal = APP.config.startDateField ? item[APP.config.startDateField] : null;
      var endVal = APP.config.endDateField ? item[APP.config.endDateField] : null;
      var startD = parseDate(startVal);
      var endD = parseDate(endVal);

      if (startD && endD) {
        var overlaps = false;
        if (APP.calendarStartDate && APP.calendarEndDate) {
          overlaps = startD <= APP.calendarEndDate && endD >= APP.calendarStartDate;
        }

        if (overlaps) {
          if (endD < today) {
            pastDueItems.push(item);
          } else {
            scheduledItems.push(item);
          }
        }
      } else {
        unscheduledItems.push(item);
      }
    });

    // Default order: ascending start date for Scheduled and Past Due
    scheduledItems.sort(function (a, b) {
      var aD = parseDate(a[APP.config.startDateField]) || new Date(0);
      var bD = parseDate(b[APP.config.startDateField]) || new Date(0);
      return aD - bD;
    });

    pastDueItems.sort(function (a, b) {
      var aD = parseDate(a[APP.config.startDateField]) || new Date(0);
      var bD = parseDate(b[APP.config.startDateField]) || new Date(0);
      return aD - bD;
    });

    // Update sidebar tabs
    var tabsContainer = el('sidebarTabs');
    if (tabsContainer) {
      clearEl(tabsContainer);
      var tabs = [
        { id: 'scheduled', label: 'To Do', count: scheduledItems.length, showCount: true },
        { id: 'unscheduled', label: 'Unscheduled', count: unscheduledItems.length, showCount: true },
        { id: 'pastDue', label: 'Past Due', count: pastDueItems.length, showCount: false },
        { id: 'archive', label: 'Archive', count: archivedItems.length, showCount: false }
      ];
      tabs.forEach(function (tab) {
        var tabText = tab.label + (tab.showCount ? ' (' + tab.count + ')' : '');
        var btn = create('button', {
          className: 'sidebar-tab' + (activeTab === tab.id ? ' active' : ''),
          textContent: tabText,
          onClick: function (e) {
            e.stopPropagation();
            APP.currentTab = tab.id;
            renderCards();
          }
        });
        tabsContainer.appendChild(btn);
      });
    }

    var itemsToRender = [];
    var tabTitles = {
      scheduled: 'To Do Items',
      pastDue: 'Past Due Items',
      unscheduled: 'Unscheduled Items',
      archive: 'Archived Items'
    };

    if (activeTab === 'scheduled') {
      itemsToRender = scheduledItems;
    } else if (activeTab === 'pastDue') {
      itemsToRender = pastDueItems;
    } else if (activeTab === 'unscheduled') {
      itemsToRender = unscheduledItems;
    } else if (activeTab === 'archive') {
      itemsToRender = archivedItems;
    }

    el('sidebarTitle').textContent = tabTitles[activeTab];
    el('sidebarCount').textContent = itemsToRender.length + ' item' + (itemsToRender.length === 1 ? '' : 's');

    if (itemsToRender.length === 0) {
      var empty = create('div', { className: 'empty-state' });
      empty.appendChild(create('div', { className: 'icon', textContent: '\uD83D\uDCCB' }));
      empty.appendChild(create('h3', { textContent: 'No items in this tab' }));
      empty.appendChild(create('p', { textContent: 'There are no items categorized under ' + tabTitles[activeTab] + '.' }));
      container.appendChild(empty);
      return;
    }

    itemsToRender.forEach(function (item) {
      var card = createCard(item);
      container.appendChild(card);
    });
  }

  function createCard(item) {
    var startVal = APP.config.startDateField ? item[APP.config.startDateField] : null;
    var endVal = APP.config.endDateField ? item[APP.config.endDateField] : null;
    var startD = parseDate(startVal);
    var endD = parseDate(endVal);
    var hasStart = !!startD;
    var hasEnd = !!endD;
    var isOnCalendar = hasStart && hasEnd;
    var hasBothDateFields = APP.config.startDateField && APP.config.endDateField;

    var card = create('div', {
      className: 'card' + (APP.selectedItemId === item.Id ? ' highlighted' : ''),
      id: 'card-' + item.Id,
      'data-item-id': String(item.Id)
    });

    var isArchived = !!(APP._archivedIds && APP._archivedIds[item.Id]);

    // Action button (top right)
    if (APP.hasContribute) {
      var actionDiv = create('div', { className: 'card-action' });
      if (isArchived) {
        var btnRestore = create('button', {
          className: 'btn btn-secondary btn-sm',
          textContent: 'Restore',
          onClick: function (e) { e.stopPropagation(); restoreItem(item.Id); }
        });
        actionDiv.appendChild(btnRestore);
      } else if (hasBothDateFields) {
        if (isOnCalendar) {
          var btnEdit = create('button', {
            className: 'btn btn-secondary btn-sm',
            textContent: 'Edit',
            onClick: function (e) { e.stopPropagation(); toggleAddToCalendar(item.Id); }
          });
          actionDiv.appendChild(btnEdit);

          var today = new Date();
          today.setHours(0, 0, 0, 0);
          var isPastDue = endD && endD < today;
          if (!isPastDue) {
            var btnRemove = create('button', {
              className: 'btn btn-danger btn-sm',
              textContent: 'Remove',
              onClick: function (e) { e.stopPropagation(); removeFromCalendar(item.Id); }
            });
            actionDiv.appendChild(btnRemove);
          }
        } else {
          var btnAdd = create('button', {
            className: 'btn btn-primary btn-sm',
            textContent: 'Add to Calendar',
            onClick: function (e) { e.stopPropagation(); toggleAddToCalendar(item.Id); }
          });
          actionDiv.appendChild(btnAdd);

          var btnArchive = create('button', {
            className: 'btn btn-secondary btn-sm',
            textContent: 'Archive',
            onClick: function (e) { e.stopPropagation(); archiveItem(item.Id); }
          });
          actionDiv.appendChild(btnArchive);
        }
      } else {
        // Missing date fields configuration, but still unscheduled so show Archive
        var btnArchive = create('button', {
          className: 'btn btn-secondary btn-sm',
          textContent: 'Archive',
          onClick: function (e) { e.stopPropagation(); archiveItem(item.Id); }
        });
        actionDiv.appendChild(btnArchive);
      }
      card.appendChild(actionDiv);
    } else if (!APP.hasContribute && hasBothDateFields) {
      // Read-only indicator
      if (isOnCalendar) {
        var badge = create('div', { className: 'card-action' });
        badge.appendChild(create('span', { className: 'btn btn-ghost btn-sm', textContent: '\uD83D\uDCC5 On Calendar', style: 'cursor:default;' }));
        card.appendChild(badge);
      }
    }

    // Display fields - sorted: Heading > Sub-heading > Regular > Label
    sortedMappings().forEach(function (mapping) {
      var val = item[mapping.internalName];
      var valStr = formatFieldValue(val);
      if (!valStr) return;
      if (mapping.style === 'heading') {
        card.appendChild(create('div', { className: 'card-field-heading', textContent: valStr }));
      } else if (mapping.style === 'subheading') {
        card.appendChild(create('div', { className: 'card-field-subheading', textContent: valStr }));
      } else if (mapping.style === 'label') {
        card.appendChild(create('span', { className: 'card-field-label', textContent: valStr }));
      } else {
        card.appendChild(create('div', { className: 'card-field-regular', textContent: valStr }));
      }
    });

    // Date display
    if (hasBothDateFields) {
      var datesDiv = create('div', { className: 'card-dates' });
      datesDiv.appendChild(create('span', { className: 'date-label', textContent: 'Start:' }));
      datesDiv.appendChild(create('span', { textContent: hasStart ? toDateStr(startVal) : '-' }));
      datesDiv.appendChild(create('span', { textContent: '|', style: 'color:var(--color-border-strong);' }));
      datesDiv.appendChild(create('span', { className: 'date-label', textContent: 'End:' }));
      datesDiv.appendChild(create('span', { textContent: hasEnd ? toDateStr(endVal) : '-' }));
      card.appendChild(datesDiv);
    }

    // Inline date picker for "Add to Calendar"
    if (APP.addingDatesForId === item.Id) {
      card.appendChild(createDateEditUI(item));
    }

    // Click handler - select this item
    card.addEventListener('click', function () { selectItem(item.Id); });

    return card;
  }

  function createDateEditUI(item) {
    var editDiv = create('div', { className: 'card-date-edit' });

    var startVal = APP.config.startDateField ? item[APP.config.startDateField] : null;
    var endVal = APP.config.endDateField ? item[APP.config.endDateField] : null;

    var startLabel = create('label', { textContent: 'Start Date' });
    var startInput = create('input', {
      type: 'date',
      id: 'datePickStart-' + item.Id,
      value: toDateStr(startVal)
    });
    editDiv.appendChild(startLabel);
    editDiv.appendChild(startInput);

    var endLabel = create('label', { textContent: 'End Date' });
    var endInput = create('input', {
      type: 'date',
      id: 'datePickEnd-' + item.Id,
      value: toDateStr(endVal)
    });
    editDiv.appendChild(endLabel);
    editDiv.appendChild(endInput);

    var actions = create('div', { className: 'date-edit-actions' });
    actions.appendChild(create('button', {
      className: 'btn btn-primary btn-sm',
      textContent: 'Save Dates',
      onClick: function (e) {
        e.stopPropagation();
        var s = startInput.value;
        var en = endInput.value;
        if (!s || !en) return;
        if (new Date(s) > new Date(en)) return;
        addToCalendar(item.Id, s, en);
      }
    }));
    actions.appendChild(create('button', {
      className: 'btn btn-ghost btn-sm',
      textContent: 'Cancel',
      onClick: function (e) {
        e.stopPropagation();
        APP.addingDatesForId = null;
        APP.selectedItemId = null;
        renderCalendar();
        renderCards();
      }
    }));
    editDiv.appendChild(actions);
    return editDiv;
  }

  function toggleAddToCalendar(itemId) {
    var startingEdit = APP.addingDatesForId !== itemId;
    APP.addingDatesForId = startingEdit ? itemId : null;
    if (startingEdit) {
      APP.selectedItemId = itemId;
    } else {
      APP.selectedItemId = null;
    }
    renderCalendar();
    renderCards();
  }

  function addToCalendar(itemId, startDate, endDate) {
    showSaving();

    var updateBody = {};
    updateBody[APP.config.startDateField] = localDateToISO(startDate);
    updateBody[APP.config.endDateField] = localDateToISO(endDate);

    var itemUrl = APP.config.listUrl + "/items(" + itemId + ")";

    spMerge(itemUrl, updateBody, APP.config.entityType)
      .then(function () {
        // Update local item
        var item = APP.items.find(function (i) { return i.Id === itemId; });
        if (item) {
          item[APP.config.startDateField] = updateBody[APP.config.startDateField];
          item[APP.config.endDateField] = updateBody[APP.config.endDateField];
        }
        APP.addingDatesForId = null;

        appendCSVRow('DATE_SET', String(itemId), APP.config.startDateField, startDate, 'Added to calendar');
        appendCSVRow('DATE_SET', String(itemId), APP.config.endDateField, endDate, 'Added to calendar');
        return saveCSV();
      })
      .then(function () {
        showSaved();
        renderCalendar();
        renderCards();
      })
      .catch(function () {
        showSaveError();
      });
  }

  function removeFromCalendar(itemId) {
    showSaving();

    var updateBody = {};
    updateBody[APP.config.startDateField] = null;
    updateBody[APP.config.endDateField] = null;

    var itemUrl = APP.config.listUrl + "/items(" + itemId + ")";

    spMerge(itemUrl, updateBody, APP.config.entityType)
      .then(function () {
        var item = APP.items.find(function (i) { return i.Id === itemId; });
        if (item) {
          item[APP.config.startDateField] = null;
          item[APP.config.endDateField] = null;
        }

        appendCSVRow('DATE_CLEARED', String(itemId), APP.config.startDateField, '', 'Removed from calendar');
        appendCSVRow('DATE_CLEARED', String(itemId), APP.config.endDateField, '', 'Removed from calendar');
        return saveCSV();
      })
      .then(function () {
        showSaved();
        renderCalendar();
        renderCards();
      })
      .catch(function () {
        showSaveError();
      });
  }

  function archiveItem(itemId) {
    showSaving();
    appendCSVRow('ARCHIVE_ITEM', String(itemId), 'Archived', 'true', '');
    saveCSV()
      .then(function () {
        if (!APP._archivedIds) APP._archivedIds = {};
        APP._archivedIds[itemId] = true;
        showSaved();
        renderCalendar();
        renderCards();
      })
      .catch(function () {
        showSaveError();
      });
  }

  function restoreItem(itemId) {
    showSaving();
    appendCSVRow('ARCHIVE_ITEM', String(itemId), 'Archived', 'false', '');
    saveCSV()
      .then(function () {
        if (APP._archivedIds) {
          delete APP._archivedIds[itemId];
        }
        showSaved();
        renderCalendar();
        renderCards();
      })
      .catch(function () {
        showSaveError();
      });
  }

  function selectItem(itemId) {
    if (APP.addingDatesForId && APP.addingDatesForId !== itemId) {
      // Cancel edit mode on other item when switching focus
      APP.addingDatesForId = null;
      renderCalendar();
      renderCards();
    }

    if (APP.addingDatesForId === itemId) {
      // Prevent toggling/collapsing if we are currently editing this item
      APP.selectedItemId = itemId;
      return;
    }

    // Toggle: if clicking same item, deselect; if different, select new
    var wasExpanded = APP.selectedItemId === itemId;
    APP.selectedItemId = wasExpanded ? null : itemId;

    // Update calendar blocks: toggle selected + expanded
    document.querySelectorAll('.calendar-block').forEach(function (b) {
      var blockId = parseInt(b.getAttribute('data-item-id'));
      var isThis = blockId === APP.selectedItemId || blockId === APP.addingDatesForId;
      b.classList.toggle('selected', isThis);
      b.classList.toggle('expanded', isThis);
    });

    // Update cards: toggle highlighted + expanded
    document.querySelectorAll('.card').forEach(function (c) {
      var cardId = parseInt(c.getAttribute('data-item-id'));
      var isThis = cardId === APP.selectedItemId || cardId === APP.addingDatesForId;
      c.classList.toggle('highlighted', isThis);
      c.classList.toggle('expanded', isThis);
    });

    // Re-stack blocks to account for height changes
    restackBlocks();

    // Scroll card into view
    if (APP.selectedItemId) {
      var card = el('card-' + APP.selectedItemId);
      if (card) {
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }
  }

  function deselectAll() {
    if (APP.selectedItemId == null && APP.addingDatesForId == null) return;
    APP.selectedItemId = null;
    APP.addingDatesForId = null;
    document.querySelectorAll('.calendar-block').forEach(function (b) {
      b.classList.remove('selected', 'expanded');
    });
    document.querySelectorAll('.card').forEach(function (c) {
      c.classList.remove('highlighted', 'expanded');
    });
    renderCalendar();
    renderCards();
    restackBlocks();
  }

  // ===== Calendar Navigation =====
  function navigateMonth(delta) {
    var start = APP.calendarStartDate || new Date();
    var nextMonth = start.getMonth() + delta;
    APP.calendarStartDate = new Date(start.getFullYear(), nextMonth, 1);
    
    if (APP.currentView === 'month') {
      APP.calendarEndDate = new Date(start.getFullYear(), nextMonth + 1, 0);
    } else {
      APP.calendarEndDate = new Date(start.getFullYear(), nextMonth + 2, 0);
    }
    renderCalendar();
    renderCards();
  }

  function goToToday() {
    var now = new Date();
    APP.calendarStartDate = new Date(now.getFullYear(), now.getMonth(), 1);
    
    if (APP.currentView === 'month') {
      APP.calendarEndDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    } else {
      APP.calendarEndDate = new Date(now.getFullYear(), now.getMonth() + 2, 0);
    }
    renderCalendar();
    renderCards();
    if (APP.currentView === 'timeline') {
      scrollToToday();
    }
  }

  // ===== Multi-User Polling =====
  function startPolling() {
    stopPolling();

    // CSV poll: every 10 seconds
    APP.pollTimerCSV = setInterval(function () {
      if (document.hidden || APP.saving) return;
      pollCSV();
    }, 10000);

    // List poll: every 60 seconds
    APP.pollTimerList = setInterval(function () {
      if (document.hidden || APP.saving) return;
      pollList();
    }, 60000);

    // Tab visibility
    document.addEventListener('visibilitychange', handleVisibility);
  }

  function stopPolling() {
    if (APP.pollTimerCSV) { clearInterval(APP.pollTimerCSV); APP.pollTimerCSV = null; }
    if (APP.pollTimerList) { clearInterval(APP.pollTimerList); APP.pollTimerList = null; }
    document.removeEventListener('visibilitychange', handleVisibility);
  }

  function handleVisibility() {
    if (document.hidden) {
      // Polling continues via the interval checks (which skip if hidden)
    } else {
      // Immediate poll on return
      pollCSV();
      pollList();
    }
  }

  function pollCSV() {
    loadCSV().then(function (rows) {
      if (rows.length !== APP.lastCSVLength) {
        // New entries detected - check for config or data changes from other users
        var hasConfigChange = false;
        for (var i = APP.lastCSVLength; i < rows.length; i++) {
          if (rows[i].ActionType === 'SETUP_CONFIG') hasConfigChange = true;
        }
        APP.lastCSVLength = rows.length;
        if (hasConfigChange) {
          replayConfig();
          loadItems().then(function () {
            renderCalendar();
            renderCards();
          });
        }
      }
    });
  }

  function pollList() {
    if (!APP.config) return;
    loadItems().then(function () {
      // Check if data actually changed
      var newFingerprint = itemsFingerprint();
      if (newFingerprint === APP._lastItemsFingerprint) return; // No changes
      APP._lastItemsFingerprint = newFingerprint;
      console.log('[SPCalendar] List data changed, check sync status');

      // Sync: check for list changes vs CSV state
      var csvChanged = false;
      APP.items.forEach(function (item) {
        var csvDateEntries = APP.csvRows.filter(function (r) {
          return r.ItemID === String(item.Id) && (r.ActionType === 'DATE_SET' || r.ActionType === 'DATE_CLEARED');
        });
        if (csvDateEntries.length > 0) {
          var lastEntry = csvDateEntries[csvDateEntries.length - 1];
          var listStart = item[APP.config.startDateField] ? toDateStr(item[APP.config.startDateField]) : '';
          var csvStart = lastEntry.Key === APP.config.startDateField ? lastEntry.Value : '';

          if (listStart && listStart !== csvStart && lastEntry.ActionType !== 'LIST_SYNC') {
            appendCSVRow('LIST_SYNC', String(item.Id), APP.config.startDateField, listStart, 'External change detected');
            csvChanged = true;
          }
        }
      });

      if (csvChanged) {
        saveCSV().catch(function () { /* silent */ });
      }

      // If user is actively typing or configuring dates, defer UI re-renders
      if (APP.addingDatesForId != null) {
        console.log('[SPCalendar] Deferring UI update since user is editing dates for item:', APP.addingDatesForId);
        return;
      }

      renderCalendar();
      renderCards();
    });
  }

  function itemsFingerprint() {
    return APP.items.map(function (item) {
      return item.Id + ':' + (item.Modified || '');
    }).join('|');
  }

  // ===== Embed Modal =====
  function showEmbedModal() {
    var modal = el('embedModal');
    modal.classList.remove('hidden');

    var pageUrl = window.location.href.split('?')[0];
    var iframeCode = '<iframe src="' + pageUrl + '?embed=true" width="100%" height="800" frameborder="0" style="border:none;"></iframe>';

    var codeBlock = el('embedCode');
    codeBlock.textContent = iframeCode;
  }

  function copyEmbedCode() {
    var code = el('embedCode').textContent;
    if (navigator.clipboard) {
      navigator.clipboard.writeText(code).then(function () {
        el('btnCopyEmbed').textContent = 'Copied!';
        setTimeout(function () { el('btnCopyEmbed').textContent = 'Copy to Clipboard'; }, 2000);
      });
    }
  }

  // ===== Initialization =====
  function init() {
    // Check embed mode
    APP.isEmbedMode = window.location.search.indexOf('embed=true') !== -1;

    // Wire up header buttons
    el('btnReconfigure').addEventListener('click', function () {
      setupStep = 0;
      hide(el('mainContainer'));
      showSetup();
    });
    el('linkReconfigureLimit').addEventListener('click', function (e) {
      e.preventDefault();
      setupStep = 0;
      hide(el('mainContainer'));
      showSetup();
    });
    el('btnCloseLimitBanner').addEventListener('click', function (e) {
      e.stopPropagation();
      // Dismiss for 14 days (2 weeks)
      var expiry = Date.now() + 14 * 24 * 60 * 60 * 1000;
      localStorage.setItem('SPCalendar_LimitDismissedUntil', String(expiry));
      hide(el('limitBanner'));
    });
    el('btnEmbed').addEventListener('click', showEmbedModal);
    el('btnCopyEmbed').addEventListener('click', copyEmbedCode);
    el('btnCloseEmbed').addEventListener('click', function () {
      el('embedModal').classList.add('hidden');
    });

    // Mobile View Switcher toggles
    el('btnViewCalendar').addEventListener('click', function () {
      el('btnViewList').classList.remove('active');
      el('btnViewCalendar').classList.add('active');
      el('mainContainer').classList.remove('mobile-list');
      el('mainContainer').classList.add('mobile-calendar');
      scrollToToday();
    });
    el('btnViewList').addEventListener('click', function () {
      el('btnViewCalendar').classList.remove('active');
      el('btnViewList').classList.add('active');
      el('mainContainer').classList.remove('mobile-calendar');
      el('mainContainer').classList.add('mobile-list');
    });

    // Calendar View Switcher toggles
    el('btnViewMonthGrid').addEventListener('click', function () { setView('month'); });
    el('btnViewTimeline').addEventListener('click', function () { setView('timeline'); });

    // Calendar navigation
    el('btnPrevMonth').addEventListener('click', function () { navigateMonth(-1); });
    el('btnNextMonth').addEventListener('click', function () { navigateMonth(1); });
    el('btnToday').addEventListener('click', goToToday);

    // Click outside to deselect
    document.addEventListener('click', function (e) {
      if (!APP.selectedItemId) return;
      var target = e.target;
      if (target.closest('.calendar-block') || target.closest('.card') || target.closest('.sidebar')) return;
      deselectAll();
    });

    // Start loading
    Promise.all([
      getCurrentUser().catch(function () { APP.currentUser = { Title: 'Unknown', Email: '' }; }),
      checkPermissions().catch(function () { APP.hasContribute = false; }),
      loadCSV()
    ]).then(function () {
      // Show permission banner if needed
      if (!APP.hasContribute) {
        show(el('permissionBanner'));
      }

      // Try to replay config from CSV
      if (replayConfig()) {
        launchApp();
      } else {
        // No config found - show setup
        if (!APP.hasContribute) {
          // Can't set up without permissions
          show(el('mainContainer'));
          var body = el('sidebarScroll');
          var empty = create('div', { className: 'empty-state' });
          empty.appendChild(create('div', { className: 'icon', textContent: '\uD83D\uDD12' }));
          empty.appendChild(create('h3', { textContent: 'Setup Required' }));
          empty.appendChild(create('p', { textContent: 'A user with Contribute permissions must complete the initial setup.' }));
          body.appendChild(empty);
        } else {
          showSetup();
        }
      }
    });
  }

  // Launch when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
</script>
</body>
</html>
