#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EI 期刊智能筛选工具 - Web 版 v2
改进：
  1. Bug修复：保留所有行包括重复标题
  2. 学科分类树：11个一级目录，勾选式
  3. 全部筛选项改为勾选/下拉多选
  4. 搜索按钮（不再实时触发）
  5. 期刊标题可点击跳转（出版商+标题检索Google Scholar）
"""

import sys, os, json, uuid, threading, webbrowser, time, copy
from collections import Counter
from datetime import datetime
from urllib.parse import quote_plus

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from flask import Flask, request, jsonify, send_file, render_template_string

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOLKIT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, TOOLKIT_DIR)

from ei_toolkit import (
    parse_ei_excel, build_column_index, apply_filter, sort_results,
    generate_stats, SUBJECT_ZH_MAP, translate_subjects, translate_titles_batch,
    COL_SOURCE_TITLE, COL_SOURCE_TYPE, COL_ISSN, COL_EISSN,
    COL_PUBLISHER, COL_COUNTRY, COL_LANGUAGE,
    TITLE_CACHE_FILE, write_output
)

app = Flask(__name__)
UPLOADS = {}
RESULTS = {}

METADATA_PATH = os.path.join(TOOLKIT_DIR, "data", "metadata.json")

def load_metadata():
    if os.path.exists(METADATA_PATH):
        with open(METADATA_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"publishers": [], "countries": [], "languages": [], "subjects": [], "subject_tree": {}}

# ============================================================
# HTML 页面
# ============================================================

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>EI 期刊智能筛选工具 v2</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif; background: #f0f2f5; color: #333; }

.navbar { background: linear-gradient(135deg, #1a237e, #283593); color: #fff; padding: 16px 32px; display: flex; align-items: center; justify-content: space-between; }
.navbar h1 { font-size: 20px; font-weight: 600; }
.navbar .version { font-size: 12px; opacity: 0.7; }

.container { max-width: 1400px; margin: 0 auto; padding: 24px; }

.card { background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); padding: 24px; margin-bottom: 20px; }
.card-title { font-size: 16px; font-weight: 600; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
.card-title .num { background: #1a237e; color: #fff; width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 14px; }

.steps { display: flex; gap: 0; margin-bottom: 24px; }
.step { flex: 1; padding: 12px; text-align: center; background: #e8eaf6; color: #999; font-size: 14px; position: relative; }
.step.active { background: #1a237e; color: #fff; }
.step.done { background: #4caf50; color: #fff; }
.step:not(:last-child)::after { content: '▶'; position: absolute; right: -8px; top: 50%; transform: translateY(-50%); color: #ccc; z-index: 1; }
.step.active:not(:last-child)::after { color: #1a237e; }
.step.done:not(:last-child)::after { color: #4caf50; }

.upload-zone { border: 2px dashed #aaa; border-radius: 8px; padding: 40px; text-align: center; cursor: pointer; transition: all 0.3s; }
.upload-zone:hover { border-color: #1a237e; background: #e8eaf6; }
.upload-zone.dragover { border-color: #1a237e; background: #e8eaf6; }
.upload-zone p { color: #666; margin-top: 8px; }
.upload-zone .icon { font-size: 48px; color: #aaa; }

.file-info { display: flex; align-items: center; gap: 12px; padding: 12px 16px; background: #e8f5e9; border-radius: 8px; margin-top: 12px; }
.file-info .filename { font-weight: 600; }
.file-info .filesize { color: #666; font-size: 13px; }

.btn { padding: 10px 24px; border: none; border-radius: 8px; font-size: 14px; font-weight: 500; cursor: pointer; transition: all 0.2s; display: inline-flex; align-items: center; gap: 6px; }
.btn-primary { background: #1a237e; color: #fff; }
.btn-primary:hover { background: #283593; }
.btn-primary:disabled { background: #bbb; cursor: not-allowed; }
.btn-success { background: #4caf50; color: #fff; }
.btn-success:hover { background: #43a047; }
.btn-outline { background: transparent; border: 1px solid #1a237e; color: #1a237e; }
.btn-outline:hover { background: #e8eaf6; }
.btn-sm { padding: 6px 14px; font-size: 13px; }

/* 筛选面板 */
.filter-section { margin-bottom: 20px; }
.filter-section-title { font-size: 14px; font-weight: 600; margin-bottom: 10px; color: #333; padding-bottom: 6px; border-bottom: 2px solid #e8eaf6; }
.filter-grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }

/* 下拉多选 */
.multi-select { position: relative; }
.ms-display { min-height: 38px; padding: 6px 32px 6px 10px; border: 1px solid #ddd; border-radius: 6px; cursor: pointer; display: flex; flex-wrap: wrap; gap: 4px; background: #fff; }
.ms-display:hover { border-color: #1a237e; }
.ms-tag { background: #e8eaf6; color: #1a237e; padding: 2px 8px; border-radius: 12px; font-size: 12px; display: flex; align-items: center; gap: 4px; }
.ms-tag .x { cursor: pointer; font-weight: bold; }
.ms-tag .x:hover { color: #f44336; }
.ms-placeholder { color: #999; font-size: 14px; line-height: 26px; }
.ms-arrow { position: absolute; right: 10px; top: 50%; transform: translateY(-50%); color: #999; pointer-events: none; }
.ms-dropdown { position: absolute; top: 100%; left: 0; right: 0; max-height: 300px; overflow-y: auto; background: #fff; border: 1px solid #ddd; border-radius: 6px; z-index: 100; box-shadow: 0 4px 12px rgba(0,0,0,0.1); display: none; }
.ms-dropdown.open { display: block; }
.ms-search { width: 100%; padding: 8px 10px; border: none; border-bottom: 1px solid #eee; font-size: 13px; outline: none; }
.ms-option { padding: 6px 10px; cursor: pointer; font-size: 13px; display: flex; align-items: center; gap: 6px; }
.ms-option:hover { background: #f5f7ff; }
.ms-option.selected { background: #e8eaf6; }
.ms-option input { margin: 0; }
.ms-footer { padding: 8px 10px; border-top: 1px solid #eee; display: flex; justify-content: space-between; font-size: 12px; }
.ms-footer a { color: #1a237e; cursor: pointer; }

/* 学科树 */
.subject-tree { border: 1px solid #e0e0e0; border-radius: 8px; max-height: 400px; overflow-y: auto; }
.st-category { border-bottom: 1px solid #f0f0f0; }
.st-cat-header { padding: 10px 14px; cursor: pointer; display: flex; align-items: center; gap: 8px; font-weight: 600; font-size: 14px; background: #fafafa; }
.st-cat-header:hover { background: #f0f4ff; }
.st-cat-header .arrow { transition: transform 0.2s; font-size: 12px; }
.st-cat-header.expanded .arrow { transform: rotate(90deg); }
.st-cat-count { margin-left: auto; font-size: 12px; color: #999; font-weight: normal; }
.st-cat-body { display: none; padding: 4px 0; }
.st-cat-body.open { display: block; }
.st-subject { padding: 4px 14px 4px 40px; display: flex; align-items: center; gap: 6px; font-size: 13px; cursor: pointer; }
.st-subject:hover { background: #f5f7ff; }
.st-subject input { margin: 0; }
.st-subject .count { margin-left: auto; color: #bbb; font-size: 12px; }

/* checkbox group */
.chk-group { display: flex; flex-wrap: wrap; gap: 8px; }
.chk-item { display: flex; align-items: center; gap: 4px; padding: 4px 12px; background: #f5f5f5; border-radius: 16px; font-size: 13px; cursor: pointer; }
.chk-item input { width: auto; }
.chk-item.checked { background: #e8eaf6; color: #1a237e; }

/* 统计卡片 */
.stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }
.stat-card { background: #fff; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
.stat-card .num { font-size: 32px; font-weight: 700; color: #1a237e; }
.stat-card .label { font-size: 13px; color: #888; margin-top: 4px; }

/* 表格 */
.table-wrapper { overflow: auto; border-radius: 8px; border: 1px solid #e0e0e0; max-height: 650px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
thead th { background: #1a237e; color: #fff; padding: 10px 12px; text-align: left; position: sticky; top: 0; white-space: nowrap; z-index: 5; }
tbody td { padding: 8px 12px; border-bottom: 1px solid #eee; vertical-align: top; }
tbody tr:hover { background: #f5f7ff; }
.title-link { color: #1a237e; text-decoration: none; cursor: pointer; }
.title-link:hover { text-decoration: underline; color: #283593; }

/* 分布图 */
.dist-row { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; }
.dist-label { width: 300px; text-align: right; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.dist-bar { height: 20px; background: #1a237e; border-radius: 4px; min-width: 2px; }
.dist-count { font-size: 13px; font-weight: 600; min-width: 40px; }

.loading { display: none; text-align: center; padding: 40px; }
.loading.show { display: block; }
.spinner { width: 48px; height: 48px; border: 4px solid #e8eaf6; border-top: 4px solid #1a237e; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 12px; }
@keyframes spin { 100% { transform: rotate(360deg); } }

.search-box { padding: 8px 14px; border: 1px solid #ddd; border-radius: 8px; width: 250px; font-size: 14px; }
.search-box:focus { outline: none; border-color: #1a237e; }

.tabs { display: flex; gap: 0; border-bottom: 2px solid #e0e0e0; margin-bottom: 16px; }
.tab { padding: 10px 24px; cursor: pointer; font-size: 14px; font-weight: 500; color: #666; border-bottom: 3px solid transparent; margin-bottom: -2px; }
.tab.active { color: #1a237e; border-bottom-color: #1a237e; }

.hidden { display: none !important; }

.toast { position: fixed; top: 20px; right: 20px; padding: 12px 24px; border-radius: 8px; color: #fff; font-size: 14px; z-index: 9999; opacity: 0; transition: opacity 0.3s; }
.toast.show { opacity: 1; }
.toast.success { background: #4caf50; }
.toast.error { background: #f44336; }
.toast.info { background: #2196f3; }

.preset-list { display: flex; flex-wrap: wrap; gap: 10px; }
.preset-card { border: 2px solid #e0e0e0; border-radius: 8px; padding: 10px 14px; cursor: pointer; transition: all 0.2s; min-width: 180px; }
.preset-card:hover { border-color: #1a237e; }
.preset-card.selected { border-color: #1a237e; background: #e8eaf6; }
.preset-card .pname { font-weight: 600; font-size: 13px; }
.preset-card .pdesc { font-size: 11px; color: #888; margin-top: 3px; }
</style>
</head>
<body>

<div class="navbar">
  <h1>EI 期刊智能筛选工具</h1>
  <span class="version">v2.0</span>
</div>

<div class="container">
  <div class="steps">
    <div class="step active" id="step1">① 上传列表</div>
    <div class="step" id="step2">② 配置筛选</div>
    <div class="step" id="step3">③ 查看结果</div>
    <div class="step" id="step4">④ 导出文件</div>
  </div>

  <!-- Step 1: 上传 -->
  <div class="card" id="card-upload">
    <div class="card-title"><span class="num">1</span> 上传 EI 列表文件</div>
    <div class="upload-zone" id="uploadZone" onclick="document.getElementById('fileInput').click()">
      <div class="icon">📁</div>
      <p><b>点击选择文件</b> 或拖拽 Excel 到此处</p>
      <p style="font-size:12px;color:#aaa;">支持 .xlsx 格式，如 CPXSourceList_072026.xlsx</p>
    </div>
    <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none">
    <div id="fileInfo" class="file-info hidden">
      <span>✅</span>
      <div>
        <div class="filename" id="fileName"></div>
        <div class="filesize" id="fileSize"></div>
      </div>
      <button class="btn btn-outline btn-sm" style="margin-left:auto;" onclick="resetUpload()">重新上传</button>
    </div>
    <div id="uploadStats" style="margin-top:16px;"></div>
  </div>

  <!-- Step 2: 筛选配置 -->
  <div class="card hidden" id="card-filter">
    <div class="card-title"><span class="num">2</span> 筛选配置</div>

    <!-- 预设方案 -->
    <div class="filter-section">
      <div class="filter-section-title">⚡ 快速预设方案</div>
      <div class="preset-list" id="presetList"></div>
    </div>

    <!-- 来源类型 -->
    <div class="filter-section">
      <div class="filter-section-title">来源类型 (Source Type)</div>
      <div class="chk-group" id="sourceTypeGroup">
        <label class="chk-item checked"><input type="checkbox" value="Journal" checked> Journal</label>
        <label class="chk-item"><input type="checkbox" value="Book-Series"> Book-Series</label>
        <label class="chk-item"><input type="checkbox" value="Trade Journal"> Trade Journal</label>
      </div>
    </div>

    <!-- 排除语言 -->
    <div class="filter-section">
      <div class="filter-section-title">排除语言（含该关键词的语言将被排除）</div>
      <div class="multi-select" id="msExcludeLang">
        <div class="ms-display" onclick="toggleDropdown('msExcludeLang')">
          <span class="ms-placeholder">点击选择要排除的语言...</span>
          <span class="ms-arrow">▼</span>
        </div>
        <div class="ms-dropdown">
          <input class="ms-search" placeholder="搜索语言..." oninput="filterDropdownOptions(this, 'msExcludeLang')">
          <div class="ms-options"></div>
          <div class="ms-footer">
            <a onclick="selectAllDropdown('msExcludeLang', true)">全选</a>
            <a onclick="selectAllDropdown('msExcludeLang', false)">清空</a>
          </div>
        </div>
      </div>
    </div>

    <!-- 出版商数量分级快速筛选 -->
    <div class="filter-section">
      <div class="filter-section-title">出版商数量分级（快速勾选对应数量区间的出版商）</div>
      <div id="publisherTierGroup" style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;"></div>
      <div style="font-size:12px;color:#888;">勾选上方选项，自动选中对应区间的出版商（例如：勾选"大型≥200"则自动选中所有 ≥200 本的出版商）</div>
    </div>

    <!-- 出版商 -->
    <div class="filter-section">
      <div class="filter-section-title">出版商 (Publisher) — 模糊匹配</div>
      <div class="multi-select" id="msPublisher">
        <div class="ms-display" onclick="toggleDropdown('msPublisher')">
          <span class="ms-placeholder">点击选择出版商（可多选）...</span>
          <span class="ms-arrow">▼</span>
        </div>
        <div class="ms-dropdown">
          <input class="ms-search" placeholder="搜索出版商..." oninput="filterDropdownOptions(this, 'msPublisher')">
          <div class="ms-options"></div>
          <div class="ms-footer">
            <span id="msPublisher_count"></span>
            <div><a onclick="selectAllDropdown('msPublisher', true)">全选</a> | <a onclick="selectAllDropdown('msPublisher', false)">清空</a></div>
          </div>
        </div>
      </div>
    </div>

    <!-- 国家/地区 -->
    <div class="filter-section">
      <div class="filter-section-title">国家/地区 (Country/Region) — 精确匹配</div>
      <div class="multi-select" id="msCountry">
        <div class="ms-display" onclick="toggleDropdown('msCountry')">
          <span class="ms-placeholder">点击选择国家/地区（可多选）...</span>
          <span class="ms-arrow">▼</span>
        </div>
        <div class="ms-dropdown">
          <input class="ms-search" placeholder="搜索国家..." oninput="filterDropdownOptions(this, 'msCountry')">
          <div class="ms-options"></div>
          <div class="ms-footer">
            <span id="msCountry_count"></span>
            <div><a onclick="selectAllDropdown('msCountry', true)">全选</a> | <a onclick="selectAllDropdown('msCountry', false)">清空</a></div>
          </div>
        </div>
      </div>
    </div>

    <!-- 语言（包含） -->
    <div class="filter-section">
      <div class="filter-section-title">语言 (Language) — 仅显示选中语言，精确匹配</div>
      <div class="multi-select" id="msLanguage">
        <div class="ms-display" onclick="toggleDropdown('msLanguage')">
          <span class="ms-placeholder">不选则不限制语言...</span>
          <span class="ms-arrow">▼</span>
        </div>
        <div class="ms-dropdown">
          <input class="ms-search" placeholder="搜索语言..." oninput="filterDropdownOptions(this, 'msLanguage')">
          <div class="ms-options"></div>
          <div class="ms-footer">
            <a onclick="selectAllDropdown('msLanguage', true)">全选</a>
            <a onclick="selectAllDropdown('msLanguage', false)">清空</a>
          </div>
        </div>
      </div>
    </div>

    <!-- 学科分类树 -->
    <div class="filter-section">
      <div class="filter-section-title">学科分类 (Subject) — 勾选学科方向，模糊匹配 Subject 1-8</div>
      <div class="subject-tree" id="subjectTree"></div>
    </div>

    <!-- 选项与按钮 -->
    <div style="display:flex;gap:24px;align-items:center;flex-wrap:wrap;margin-top:20px;padding-top:16px;border-top:2px solid #e8eaf6;">
      <label style="font-size:14px;display:flex;align-items:center;gap:6px;">
        <input type="checkbox" id="optTranslate" checked> 中英对照翻译
      </label>
      <button class="btn btn-primary" id="btnFilter" onclick="runFilter()" style="margin-left:auto;">
        🔍 执行筛选
      </button>
    </div>

    <div class="loading" id="filterLoading">
      <div class="spinner"></div>
      <p id="filterLoadingText">正在筛选...</p>
    </div>
  </div>

  <!-- Step 3: 结果 -->
  <div class="card hidden" id="card-result">
    <div class="card-title">
      <span class="num">3</span> 筛选结果
      <span style="margin-left:auto;display:flex;gap:8px;">
        <input type="text" class="search-box" id="tableSearch" placeholder="🔍 输入关键词后点击搜索...">
        <button class="btn btn-outline btn-sm" onclick="filterTable()">搜索</button>
        <button class="btn btn-success btn-sm" onclick="exportExcel()">📥 导出 Excel</button>
      </span>
    </div>

    <div class="stats-grid" id="statsGrid"></div>

    <div class="tabs">
      <div class="tab active" onclick="switchTab(event,'table')">数据表格</div>
      <div class="tab" onclick="switchTab(event,'dist')">分布统计</div>
    </div>

    <div id="view-table">
      <div class="table-wrapper">
        <table id="resultTable">
          <thead><tr id="tableHead"></tr></thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
      <div style="margin-top:8px;font-size:13px;color:#888;" id="tableInfo"></div>
    </div>

    <div id="view-dist" class="hidden">
      <div class="filter-grid2">
        <div><h3 style="margin-bottom:12px;">出版社分布 Top 20</h3><div id="distPublisher"></div></div>
        <div><h3 style="margin-bottom:12px;">国家/地区分布 Top 20</h3><div id="distCountry"></div></div>
        <div><h3 style="margin-bottom:12px;">学科分布 Top 20</h3><div id="distSubject"></div></div>
        <div><h3 style="margin-bottom:12px;">语言分布</h3><div id="distLanguage"></div></div>
      </div>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
let sessionId = null;
let resultData = null;
let metadata = null;

// ============================================================
// Toast
// ============================================================
function showToast(msg, type='info') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show ' + type;
  setTimeout(() => t.className = 'toast', 3000);
}

// ============================================================
// 下拉多选组件
// ============================================================
function initMultiSelect(id, options, withCount=false) {
  const container = document.getElementById(id);
  const optionsContainer = container.querySelector('.ms-options');
  optionsContainer.innerHTML = '';

  options.forEach(opt => {
    const val = typeof opt === 'string' ? opt : opt.value;
    const label = typeof opt === 'string' ? opt : (opt.label || opt.value);
    const count = typeof opt === 'object' ? (opt.count || '') : '';

    const div = document.createElement('div');
    div.className = 'ms-option';
    div.dataset.value = val;
    div.dataset.label = label.toLowerCase();
    div.innerHTML = `<input type="checkbox" value="${val}" onchange="updateMultiSelectDisplay('${id}')">
      <span>${label}</span>${count ? `<span class="count" style="margin-left:auto;color:#bbb;font-size:12px;">(${count})</span>` : ''}`;
    optionsContainer.appendChild(div);
  });
}

function toggleDropdown(id) {
  const dd = document.querySelector(`#${id} .ms-dropdown`);
  const isOpen = dd.classList.contains('open');
  // 关闭所有其他下拉
  document.querySelectorAll('.ms-dropdown.open').forEach(d => d.classList.remove('open'));
  if (!isOpen) dd.classList.add('open');
}

function filterDropdownOptions(input, id) {
  const q = input.value.toLowerCase();
  const options = document.querySelectorAll(`#${id} .ms-option`);
  options.forEach(opt => {
    opt.style.display = opt.dataset.label.includes(q) ? '' : 'none';
  });
}

function selectAllDropdown(id, selectAll) {
  const checkboxes = document.querySelectorAll(`#${id} .ms-option input[type="checkbox"]`);
  checkboxes.forEach(cb => {
    // 只勾选可见的
    const opt = cb.closest('.ms-option');
    if (opt.style.display !== 'none') {
      cb.checked = selectAll;
    }
  });
  updateMultiSelectDisplay(id);
}

function updateMultiSelectDisplay(id) {
  const container = document.getElementById(id);
  const display = container.querySelector('.ms-display');
  const checked = container.querySelectorAll('.ms-option input:checked');

  // 重建标签
  display.innerHTML = '';
  if (checked.length === 0) {
    const placeholders = {
      'msExcludeLang': '点击选择要排除的语言...',
      'msPublisher': '点击选择出版商（可多选）...',
      'msCountry': '点击选择国家/地区（可多选）...',
      'msLanguage': '不选则不限制语言...',
    };
    display.innerHTML = `<span class="ms-placeholder">${placeholders[id] || '点击选择...'}</span>`;
  } else {
    checked.forEach(cb => {
      const label = cb.nextElementSibling.textContent;
      const tag = document.createElement('span');
      tag.className = 'ms-tag';
      tag.innerHTML = `${label} <span class="x" onclick="event.stopPropagation();removeTag('${id}','${cb.value}')">✕</span>`;
      display.appendChild(tag);
    });
  }
  const arrow = document.createElement('span');
  arrow.className = 'ms-arrow';
  arrow.textContent = '▼';
  display.appendChild(arrow);

  // 更新计数
  const countEl = container.querySelector('.ms-footer span[id$="_count"]');
  if (countEl) {
    countEl.textContent = `已选 ${checked.length} 项`;
  }
}

function removeTag(id, value) {
  const cb = document.querySelector(`#${id} .ms-option input[value="${value}"]`);
  if (cb) {
    cb.checked = false;
    updateMultiSelectDisplay(id);
  }
}

function getSelectedValues(id) {
  const checkboxes = document.querySelectorAll(`#${id} .ms-option input:checked`);
  return Array.from(checkboxes).map(cb => cb.value);
}

// 关闭下拉当点击外部
document.addEventListener('click', (e) => {
  if (!e.target.closest('.multi-select')) {
    document.querySelectorAll('.ms-dropdown.open').forEach(d => d.classList.remove('open'));
  }
});

// ============================================================
// 出版商数量分级
// ============================================================
function initPublisherTiers(tiers) {
  const container = document.getElementById('publisherTierGroup');
  container.innerHTML = '';
  const tierLabels = {
    '≥200': '大型（≥200本）',
    '100-199': '大型（100-199本）',
    '50-99': '中型（50-99本）',
    '20-49': '中小型（20-49本）',
    '10-19': '小型（10-19本）',
    '5-9': '微型（5-9本）',
    '2-4': '超微型（2-4本）',
    '1': '单本（仅1本）'
  };
  const tierOrder = ['≥200', '100-199', '50-99', '20-49', '10-19', '5-9', '2-4', '1'];
  tierOrder.forEach(key => {
    const publishers = tiers[key] || [];
    if (publishers.length === 0) return;
    const label = document.createElement('label');
    label.className = 'chk-item';
    label.innerHTML = `<input type="checkbox" onchange="applyPublisherTier('${key}', this.checked)"> ${tierLabels[key]} (${publishers.length}家)`;
    container.appendChild(label);
  });
}

function applyPublisherTier(tierKey, checked) {
  if (!metadata || !metadata.publisher_tiers) return;
  const tierLabels = {
    '≥200': '大型（≥200本）',
    '100-199': '大型（100-199本）',
    '50-99': '中型（50-99本）',
    '20-49': '中小型（20-49本）',
    '10-19': '小型（10-19本）',
    '5-9': '微型（5-9本）',
    '2-4': '超微型（2-4本）',
    '1': '单本（仅1本）'
  };
  const pubs = metadata.publisher_tiers[tierKey] || [];
  const msOptions = document.querySelectorAll('#msPublisher .ms-option input[type="checkbox"]');
  msOptions.forEach(cb => {
    if (pubs.includes(cb.value)) {
      cb.checked = checked;
    }
  });
  updateMultiSelectDisplay('msPublisher');
  showToast(`${checked ? '已勾选' : '已取消'} ${tierLabels[tierKey]} (${pubs.length}家出版商)`, 'info');
}

// ============================================================
// 学科分类树
// ============================================================
function initSubjectTree(tree) {
  const container = document.getElementById('subjectTree');
  container.innerHTML = '';

  Object.entries(tree).forEach(([catName, catData]) => {
    const catDiv = document.createElement('div');
    catDiv.className = 'st-category';

    const header = document.createElement('div');
    header.className = 'st-cat-header';
    header.onclick = () => {
      header.classList.toggle('expanded');
      body.classList.toggle('open');
    };
    header.innerHTML = `<span class="arrow">▶</span> ${catData.icon} ${catName}
      <span class="st-cat-count">${catData.subjects.length} 个学科</span>`;

    const body = document.createElement('div');
    body.className = 'st-cat-body';

    catData.subjects.forEach(subjName => {
      const subjDiv = document.createElement('div');
      subjDiv.className = 'st-subject';
      const subjMeta = metadata.subjects.find(s => s.name === subjName);
      const count = subjMeta ? subjMeta.count : '';
      subjDiv.innerHTML = `
        <input type="checkbox" value="${subjName}" id="subj_${subjName.replace(/[^a-zA-Z0-9]/g,'_')}">
        <label for="subj_${subjName.replace(/[^a-zA-Z0-9]/g,'_')}" style="cursor:pointer;flex:1;">${subjName}</label>
        ${count ? `<span class="count">${count}</span>` : ''}`;
      body.appendChild(subjDiv);
    });

    catDiv.appendChild(header);
    catDiv.appendChild(body);
    container.appendChild(catDiv);
  });
}

function getSelectedSubjects() {
  const checkboxes = document.querySelectorAll('#subjectTree input:checked');
  return Array.from(checkboxes).map(cb => cb.value);
}

// ============================================================
// checkbox 样式
// ============================================================
document.querySelectorAll('#sourceTypeGroup input').forEach(cb => {
  cb.addEventListener('change', () => {
    cb.parentElement.classList.toggle('checked', cb.checked);
  });
});

// ============================================================
// 文件上传
// ============================================================
const uploadZone = document.getElementById('uploadZone');
const fileInput = document.getElementById('fileInput');

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length > 0) uploadFile(e.target.files[0]);
});

uploadZone.addEventListener('dragover', (e) => { e.preventDefault(); uploadZone.classList.add('dragover'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
uploadZone.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadZone.classList.remove('dragover');
  if (e.dataTransfer.files.length > 0) uploadFile(e.dataTransfer.files[0]);
});

function uploadFile(file) {
  const formData = new FormData();
  formData.append('file', file);

  document.getElementById('uploadZone').classList.add('hidden');
  document.getElementById('fileInfo').classList.remove('hidden');
  document.getElementById('fileName').textContent = file.name;
  document.getElementById('fileSize').textContent = (file.size/1024/1024).toFixed(2) + ' MB';

  fetch('/api/upload', { method: 'POST', body: formData })
    .then(r => r.json())
    .then(data => {
      if (data.error) { showToast(data.error, 'error'); resetUpload(); return; }
      sessionId = data.session_id;
      let html = '<div style="display:flex;gap:16px;flex-wrap:wrap;">';
      data.stats.forEach(s => {
        html += `<div style="background:#f5f5f5;padding:8px 16px;border-radius:8px;text-align:center;">
          <div style="font-size:20px;font-weight:700;color:#1a237e;">${s.value}</div>
          <div style="font-size:12px;color:#888;">${s.label}</div></div>`;
      });
      html += '</div>';
      document.getElementById('uploadStats').innerHTML = html;
      document.getElementById('card-filter').classList.remove('hidden');
      document.getElementById('step1').classList.add('done');
      document.getElementById('step2').classList.add('active');
      loadPresets();
      showToast('上传成功！', 'success');
    })
    .catch(err => { showToast('上传失败: ' + err, 'error'); resetUpload(); });
}

function resetUpload() {
  sessionId = null;
  fileInput.value = '';
  document.getElementById('uploadZone').classList.remove('hidden');
  document.getElementById('fileInfo').classList.add('hidden');
  document.getElementById('uploadStats').innerHTML = '';
  document.getElementById('card-filter').classList.add('hidden');
  document.getElementById('card-result').classList.add('hidden');
  document.getElementById('step1').classList.remove('done');
  document.getElementById('step2').classList.remove('active');
}

// ============================================================
// 预设
// ============================================================
function loadPresets() {
  fetch('/api/presets').then(r => r.json()).then(data => {
    const list = document.getElementById('presetList');
    list.innerHTML = '';
    data.presets.forEach(p => {
      const card = document.createElement('div');
      card.className = 'preset-card';
      card.innerHTML = `<div class="pname">${p.name}</div><div class="pdesc">${p.desc || ''}</div>`;
      card.onclick = () => {
        document.querySelectorAll('.preset-card').forEach(c => c.classList.remove('selected'));
        card.classList.add('selected');
        applyPreset(p);
      };
      list.appendChild(card);
    });
  });
}

function applyPreset(preset) {
  fetch('/api/preset_detail?file=' + encodeURIComponent(preset.file))
    .then(r => r.json())
    .then(cfg => {
      if (cfg.error) return;
      // Source type
      const sts = cfg.base_filter?.source_type || ['Journal'];
      document.querySelectorAll('#sourceTypeGroup input').forEach(cb => {
        cb.checked = sts.includes(cb.value);
        cb.parentElement.classList.toggle('checked', cb.checked);
      });
      // Exclude language
      const excludeLangs = cfg.base_filter?.exclude_language_keywords || [];
      setMultiSelectValues('msExcludeLang', excludeLangs);
      // Dimensional
      const dims = cfg.dimensional_filter || {};
      setMultiSelectValues('msPublisher', dims.publisher?.values || []);
      setMultiSelectValues('msCountry', dims.country_region?.values || []);
      setMultiSelectValues('msLanguage', dims.language?.values || []);
      // Subject
      const subjVals = dims.subject_all?.values || [];
      document.querySelectorAll('#subjectTree input').forEach(cb => {
        cb.checked = subjVals.some(v => cb.value.toLowerCase().includes(v.toLowerCase()));
      });
      showToast('已应用预设: ' + preset.name, 'info');
    });
}

function setMultiSelectValues(id, values) {
  document.querySelectorAll(`#${id} .ms-option input`).forEach(cb => {
    cb.checked = values.some(v => v.toLowerCase() === cb.value.toLowerCase());
  });
  updateMultiSelectDisplay(id);
}

// ============================================================
// 执行筛选
// ============================================================
function runFilter() {
  if (!sessionId) return;
  const btn = document.getElementById('btnFilter');
  btn.disabled = true;
  document.getElementById('filterLoading').classList.add('show');
  const doTranslate = document.getElementById('optTranslate').checked;
  document.getElementById('filterLoadingText').textContent = doTranslate ? '正在筛选 + 翻译中...' : '正在筛选...';

  const config = buildConfig();

  fetch('/api/filter', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ session_id: sessionId, config: config })
  })
  .then(r => r.json())
  .then(data => {
    if (data.error) { showToast(data.error, 'error'); return; }
    resultData = data;
    renderResults(data);
    document.getElementById('card-result').classList.remove('hidden');
    document.getElementById('step2').classList.add('done');
    document.getElementById('step3').classList.add('active');
    document.getElementById('step4').classList.add('active');
    showToast(`筛选完成: ${data.total} 条结果`, 'success');
  })
  .catch(err => showToast('筛选失败: ' + err, 'error'))
  .finally(() => {
    btn.disabled = false;
    document.getElementById('filterLoading').classList.remove('show');
  });
}

function buildConfig() {
  const sourceTypes = [];
  document.querySelectorAll('#sourceTypeGroup input:checked').forEach(cb => sourceTypes.push(cb.value));
  const excludeLangs = getSelectedValues('msExcludeLang');
  const publishers = getSelectedValues('msPublisher');
  const countries = getSelectedValues('msCountry');
  const languages = getSelectedValues('msLanguage');
  const subjects = getSelectedSubjects();

  const config = {
    base_filter: {
      source_type: sourceTypes.length ? sourceTypes : ['Journal'],
      exclude_language_keywords: excludeLangs,
    },
    dimensional_filter: {},
    sorting: {
      primary: { field: 'Publisher', order: 'desc', sort_by: 'count' },
      secondary: { field: 'Source title', order: 'asc', sort_by: 'alphabetical' }
    },
    output: { bilingual: document.getElementById('optTranslate').checked }
  };

  if (publishers.length) config.dimensional_filter.publisher = { mode: 'include', match_type: 'fuzzy', values: publishers };
  if (countries.length) config.dimensional_filter.country_region = { mode: 'include', match_type: 'exact', values: countries };
  if (languages.length) config.dimensional_filter.language = { mode: 'include', match_type: 'exact', values: languages };
  if (subjects.length) config.dimensional_filter.subject_all = { mode: 'include', match_type: 'fuzzy', values: subjects };

  return config;
}

// ============================================================
// 渲染结果
// ============================================================
function renderResults(data) {
  const grid = document.getElementById('statsGrid');
  grid.innerHTML = `
    <div class="stat-card"><div class="num">${data.stats.total_input}</div><div class="label">输入总数</div></div>
    <div class="stat-card"><div class="num">${data.stats.after_base}</div><div class="label">基础筛选后</div></div>
    <div class="stat-card"><div class="num" style="color:#4caf50;">${data.total}</div><div class="label">最终结果</div></div>
    <div class="stat-card"><div class="num">${data.stats.publishers}</div><div class="label">出版社数</div></div>
  `;
  renderTable(data);
  renderDist(data);
  document.getElementById('tableInfo').textContent =
    `共 ${data.total} 条 | 显示前 ${Math.min(data.total, 200)} 条（导出包含全部）`;
}

function renderTable(data) {
  const head = document.getElementById('tableHead');
  const body = document.getElementById('tableBody');
  head.innerHTML = '';
  body.innerHTML = '';

  data.columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    head.appendChild(th);
  });

  const rows = data.rows.slice(0, 200);
  rows.forEach(row => {
    const tr = document.createElement('tr');
    data.columns.forEach((col, i) => {
      const td = document.createElement('td');
      let val = row[i] || '';
      val = String(val);

      // Source title 列：添加可点击跳转链接
      if (col === 'Source title') {
        // 提取纯英文标题（去掉翻译部分）
        let title = val.split('\n')[0];
        let publisher = '';
        // 找 Publisher 列的值
        const pubIdx = data.columns.indexOf('Publisher');
        if (pubIdx >= 0 && row[pubIdx]) {
          publisher = String(row[pubIdx]).split('\n')[0];
        }
        // 构建搜索词：出版商 + 期刊名
        let searchTerm = title;
        if (publisher) {
          searchTerm = publisher + ' ' + title;
        }
        const searchUrl = 'https://www.google.com/search?q=' + encodeURIComponent(searchTerm);
        // 显示翻译后的完整文本（如果有）
        let displayText = val.replace(/\n/g, '<br>');
        td.innerHTML = `<a class="title-link" href="${searchUrl}" target="_blank" title="点击搜索: ${searchTerm}">${displayText}</a>`;
      } else {
        td.innerHTML = val.replace(/\n/g, '<br>');
      }
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function renderDist(data) {
  function renderBars(arr, targetId) {
    const maxVal = Math.max(...arr.map(d => d[1]));
    const el = document.getElementById(targetId);
    el.innerHTML = arr.map(item => {
      const [name, count] = item;
      const pct = (count / maxVal * 100).toFixed(1);
      return `<div class="dist-row">
        <div class="dist-label" title="${name}">${name}</div>
        <div class="dist-bar" style="width:${pct}%"></div>
        <div class="dist-count">${count}</div>
      </div>`;
    }).join('');
  }
  renderBars(data.dist.publisher, 'distPublisher');
  renderBars(data.dist.country, 'distCountry');
  renderBars(data.dist.subject, 'distSubject');
  renderBars(data.dist.language, 'distLanguage');
}

function switchTab(e, tab) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  e.currentTarget.classList.add('active');
  document.getElementById('view-table').classList.toggle('hidden', tab !== 'table');
  document.getElementById('view-dist').classList.toggle('hidden', tab !== 'dist');
}

function filterTable() {
  const q = document.getElementById('tableSearch').value.toLowerCase();
  const rows = document.querySelectorAll('#tableBody tr');
  let visible = 0;
  rows.forEach(tr => {
    const text = tr.textContent.toLowerCase();
    const show = !q || text.includes(q);
    tr.style.display = show ? '' : 'none';
    if (show) visible++;
  });
  document.getElementById('tableInfo').textContent =
    `共 ${resultData.total} 条 | 显示 ${visible} 条` + (q ? '（搜索过滤）' : '');
}

function exportExcel() {
  if (!sessionId) return;
  showToast('正在生成 Excel...', 'info');
  window.location.href = '/api/export?session_id=' + sessionId;
}

// ============================================================
// 初始化
// ============================================================
fetch('/api/metadata').then(r => r.json()).then(data => {
  metadata = data;
  // 初始化下拉多选
  initMultiSelect('msExcludeLang', data.languages);
  initMultiSelect('msPublisher', data.publishers);
  initMultiSelect('msCountry', data.countries);
  initMultiSelect('msLanguage', data.languages);
  // 初始化学科树
  initSubjectTree(data.subject_tree);
  // 初始化出版商数量分级
  initPublisherTiers(data.publisher_tiers);
});
</script>
</body>
</html>
"""


# ============================================================
# API 路由
# ============================================================

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)


@app.route('/api/metadata')
def api_metadata():
    meta = load_metadata()
    # 构建出版商数量分级
    tiers = {"≥200": [], "100-199": [], "50-99": [], "20-49": [], "10-19": [], "5-9": [], "2-4": [], "1": []}
    for pub in meta.get("publishers", []):
        c = pub.get("count", 0)
        val = pub["value"]
        if c >= 200: tiers["≥200"].append(val)
        elif c >= 100: tiers["100-199"].append(val)
        elif c >= 50: tiers["50-99"].append(val)
        elif c >= 20: tiers["20-49"].append(val)
        elif c >= 10: tiers["10-19"].append(val)
        elif c >= 5: tiers["5-9"].append(val)
        elif c >= 2: tiers["2-4"].append(val)
        else: tiers["1"].append(val)
    meta["publisher_tiers"] = tiers
    return jsonify(meta)


@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'file' not in request.files:
        return jsonify({"error": "未找到文件"}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "文件名为空"}), 400

    session_id = str(uuid.uuid4())[:8]
    upload_dir = os.path.join(TOOLKIT_DIR, "data", "uploads")
    os.makedirs(upload_dir, exist_ok=True)
    filepath = os.path.join(upload_dir, f"{session_id}_{file.filename}")
    file.save(filepath)

    try:
        header, data_rows = parse_ei_excel(filepath)
        col_map = build_column_index(header)

        st_counter = Counter()
        for row in data_rows:
            st = str(row[COL_SOURCE_TYPE]).strip() if row[COL_SOURCE_TYPE] else "Unknown"
            st_counter[st] += 1

        stats = [
            {"label": "总条目", "value": len(data_rows)},
            {"label": "Journal", "value": st_counter.get("Journal", 0)},
            {"label": "Book-Series", "value": st_counter.get("Book-Series", 0)},
            {"label": "Trade Journal", "value": st_counter.get("Trade Journal", 0)},
        ]

        UPLOADS[session_id] = {
            "filepath": filepath, "header": header, "data_rows": data_rows,
            "col_map": col_map, "filename": file.filename,
        }
        return jsonify({"session_id": session_id, "stats": stats})
    except Exception as e:
        return jsonify({"error": f"解析失败: {str(e)}"}), 500


@app.route('/api/presets')
def api_presets():
    preset_dir = os.path.join(TOOLKIT_DIR, "presets")
    presets = []
    if os.path.isdir(preset_dir):
        for fname in sorted(os.listdir(preset_dir)):
            if fname.endswith(".json") and not fname.startswith("researcher"):
                try:
                    with open(os.path.join(preset_dir, fname), "r", encoding="utf-8") as f:
                        data = json.load(f)
                    presets.append({
                        "file": fname,
                        "name": data.get("preset_name", fname),
                        "desc": data.get("preset_description", ""),
                    })
                except Exception:
                    pass
    return jsonify({"presets": presets})


@app.route('/api/preset_detail')
def api_preset_detail():
    fname = request.args.get("file", "")
    preset_path = os.path.join(TOOLKIT_DIR, "presets", fname)
    if not os.path.exists(preset_path):
        return jsonify({"error": "预设不存在"}), 404
    with open(preset_path, "r", encoding="utf-8") as f:
        return jsonify(json.load(f))


@app.route('/api/filter', methods=['POST'])
def api_filter():
    data = request.json
    session_id = data.get("session_id")
    config = data.get("config")

    if session_id not in UPLOADS:
        return jsonify({"error": "会话不存在，请重新上传"}), 400

    upload = UPLOADS[session_id]
    data_rows = copy.deepcopy(upload["data_rows"])
    col_map = upload["col_map"]

    filtered, base_stats = apply_filter(data_rows, col_map, config)
    sorting_config = config.get("sorting", {})
    sort_results(filtered, col_map, sorting_config)

    do_translate = config.get("output", {}).get("bilingual", True)
    if do_translate:
        filtered = [translate_subjects(row, col_map, SUBJECT_ZH_MAP) for row in filtered]
        all_titles = [str(row[COL_SOURCE_TITLE]).strip() if row[COL_SOURCE_TITLE] else "" for row in filtered]
        title_translations = translate_titles_batch(all_titles)
        for row in filtered:
            title = str(row[COL_SOURCE_TITLE]).strip() if row[COL_SOURCE_TITLE] else ""
            zh = title_translations.get(title)
            if zh:
                row[COL_SOURCE_TITLE] = f"{title}\n{zh}"

    dist_stats = generate_stats(filtered, col_map)
    pub_counter = Counter()
    for row in filtered:
        pub = str(row[COL_PUBLISHER]).strip() if row[COL_PUBLISHER] else "Unknown"
        pub_counter[pub] += 1

    columns = upload["header"]
    rows_json = [[str(v) if v is not None else "" for v in row] for row in filtered]

    dist = {
        "publisher": [[k, v] for k, v in pub_counter.most_common(20)],
        "country": [[k, v] for k, v in dist_stats["countries"].most_common(20)],
        "subject": [[k.split("\n")[0] if "\n" in k else k, v]
                     for k, v in dist_stats["subjects"].most_common(20)],
        "language": [[k, v] for k, v in dist_stats["languages"].most_common(20)],
    }

    RESULTS[session_id] = {
        "filtered": filtered, "columns": columns, "config": config,
        "base_stats": base_stats, "dist_stats": dist_stats,
        "pub_counter": pub_counter, "filename": upload["filename"],
    }

    return jsonify({
        "total": len(filtered),
        "columns": columns,
        "rows": rows_json,
        "stats": {
            "total_input": base_stats["total_input"],
            "after_base": base_stats["after_base_filter"],
            "publishers": len(pub_counter),
        },
        "dist": dist,
    })


@app.route('/api/export')
def api_export():
    session_id = request.args.get("session_id")
    if session_id not in RESULTS:
        return jsonify({"error": "无结果可导出"}), 400

    result = RESULTS[session_id]
    filtered = result["filtered"]
    columns = result["columns"]
    config = result["config"]
    base_stats = result["base_stats"]
    dist_stats = result["dist_stats"]
    pub_counter = result["pub_counter"]

    output_dir = os.path.join(TOOLKIT_DIR, "data", "output")
    os.makedirs(output_dir, exist_ok=True)
    preset_name = config.get("preset_name", "filtered")
    safe_name = "".join(c for c in preset_name if c not in r'\/:*?"<>|')
    output_path = os.path.join(output_dir, f"EI_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    filter_desc = {
        "source_file": result["filename"],
        "preset_name": config.get("preset_name", "自定义"),
        "base_filter_desc": "; ".join([
            f"Source type={','.join(config.get('base_filter',{}).get('source_type',[]))}",
            f"排除{','.join(config.get('base_filter',{}).get('exclude_language_keywords',[]))}"
        ]),
        "dim_filter_desc": "见筛选配置",
        "sort_desc": "按出版商数量降序",
        "translated": config.get("output", {}).get("bilingual", True),
    }

    write_output(output_path, columns, filtered, base_stats, dist_stats, filter_desc, pub_counter)
    return send_file(output_path, as_attachment=True, download_name=os.path.basename(output_path))


# ============================================================
# 启动
# ============================================================

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:5000")

if __name__ == '__main__':
    print("=" * 50)
    print("  EI 期刊智能筛选工具 - Web v2.0")
    print("  正在启动... 浏览器将自动打开")
    print("  如未自动打开，请访问: http://127.0.0.1:5000")
    print("=" * 50)
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='127.0.0.1', port=5000, debug=False)
