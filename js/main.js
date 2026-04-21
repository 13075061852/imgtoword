
    const REFERENCE_DATA = [
      {
        "型号": "320G6-B1S",
        "批次": "B506202",
        "测试温度": "260°",
        "熔指": [25.6, 25.7, 25.6],
        "拉伸强度[Mpa]": [135, 134, 136],
        "断裂伸长率[%]": [1.24, 1.28, 1.29],
        "弯曲强度[Mpa]": [207, 208, 209],
        "弯曲模量[Mpa]": [9392, 9450, 9650],
        "冲击强度[Mpa]": [7.4, 7.5, 7.6, 7.4, 75]
      },
      {
        "型号": "320G2-N1",
        "批次": "B506202",
        "测试温度": "260°",
        "熔指": [25.6, 25.7, 25.6],
        "拉伸强度[Mpa]": [135, 134, 136],
        "断裂伸长率[%]": [1.24, 1.28, 1.29],
        "弯曲强度[Mpa]": [207, 208, 209],
        "弯曲模量[Mpa]": [9392, 9450, 9650],
        "冲击强度[Mpa]": [7.4, 7.5, 7.6, 7.4, 75]
      },
      {
        "型号": "320G1-B2H",
        "批次": "B506202",
        "测试温度": "260°",
        "熔指": [25.6, 25.7, 25.6],
        "拉伸强度[Mpa]": [135, 134, 136],
        "断裂伸长率[%]": [1.24, 1.28, 1.29],
        "弯曲强度[Mpa]": [207, 208, 209],
        "弯曲模量[Mpa]": [9392, 9450, 9650],
        "冲击强度[Mpa]": [7.4, 7.5, 7.6, 7.4, 75]
      }
    ];

    const COLUMN_DEFS = [
      { key: "型号", kind: "basic", className: "group-basic" },
      { key: "批次", kind: "basic", className: "group-basic" },
      { key: "测试温度", kind: "basic", className: "group-basic" },
      { key: "熔指", kind: "measure", className: "group-melt" },
      { key: "拉伸强度[Mpa]", kind: "measure", className: "group-tensile" },
      { key: "断裂伸长率[%]", kind: "measure", className: "group-elongation" },
      { key: "弯曲强度[Mpa]", kind: "measure", className: "group-flexural" },
      { key: "弯曲模量[Mpa]", kind: "measure", className: "group-modulus" },
      { key: "冲击强度[Mpa]", kind: "measure", className: "group-impact" }
    ];

    const STORAGE_KEYS = {
      baseUrl: "lmstudio-doc-workbench-base-url",
      model: "lmstudio-doc-workbench-model",
      provider: "lmstudio-doc-workbench-provider",
      openRouterApiKey: "lmstudio-doc-workbench-openrouter-api-key",
      selectedModelByProvider: "lmstudio-doc-workbench-selected-model-by-provider",
      recognitionSnapshot: "lmstudio-doc-workbench-recognition-snapshot"
    };

    const BASE_COMPARE_ZOOM = 1.1;

    const FIELD_ALIASES = {
      "牌号": "型号", "样品型号": "型号", "产品型号": "型号",
      "批号": "批次", "批次号": "批次",
      "试验温度": "测试温度", "温度": "测试温度", "测试温度℃": "测试温度", "测试温度°C": "测试温度",
      "熔融指数": "熔指", "熔体流动速率": "熔指", "熔体流动指数": "熔指", "MFR": "熔指", "MI": "熔指",
      "拉伸强度": "拉伸强度[Mpa]", "抗拉强度": "拉伸强度[Mpa]",
      "断裂伸长率": "断裂伸长率[%]", "伸长率": "断裂伸长率[%]",
      "弯曲强度": "弯曲强度[Mpa]", "挠曲强度": "弯曲强度[Mpa]",
      "弯曲模量": "弯曲模量[Mpa]", "挠曲模量": "弯曲模量[Mpa]",
      "冲击强度": "冲击强度[Mpa]", "缺口冲击[Mpa]": "冲击强度[Mpa]",
      "缺口冲击": "冲击强度[Mpa]", "缺口冲击强度[Mpa]": "冲击强度[Mpa]",
      "缺口冲击强度": "冲击强度[Mpa]", "简支梁缺口冲击强度": "冲击强度[Mpa]",
      "冲击值[Mpa]": "冲击强度[Mpa]", "冲击值": "冲击强度[Mpa]"
    };

    function readJsonStorage(key, fallback) {
      try {
        const value = localStorage.getItem(key);
        return value ? JSON.parse(value) : fallback;
      } catch {
        return fallback;
      }
    }

    const state = {
      baseUrl: localStorage.getItem(STORAGE_KEYS.baseUrl) || "http://localhost:1234",
      provider: localStorage.getItem(STORAGE_KEYS.provider) || "lmstudio",
      openRouterApiKey: localStorage.getItem(STORAGE_KEYS.openRouterApiKey) || "",
      selectedModelByProvider: readJsonStorage(STORAGE_KEYS.selectedModelByProvider, {}),
      model: "",
      imageDataUrl: "",
      parsedRecords: [],
      rawText: "",
      connected: false,
      compareView: {
        zoom: BASE_COMPARE_ZOOM,
        fitScale: 1,
        offsetX: 0,
        offsetY: 0,
        dragging: false,
        startX: 0,
        startY: 0,
        baseX: 0,
        baseY: 0
      }
    };

    const els = {
      providerSelect: document.getElementById("providerSelect"),
      baseUrlInput: document.getElementById("baseUrlInput"),
      openRouterKeyInput: document.getElementById("openRouterKeyInput"),
      modelSelect: document.getElementById("modelSelect"),
      modelSelectRoot: document.getElementById("modelSelectRoot"),
      modelSelectTrigger: document.getElementById("modelSelectTrigger"),
      modelSelectLabel: document.getElementById("modelSelectLabel"),
      modelSelectMenu: document.getElementById("modelSelectMenu"),
      checkConnectionBtn: document.getElementById("checkConnectionBtn"),
      connectionStatusBtn: document.getElementById("connectionStatusBtn"),
      imageInput: document.getElementById("imageInput"),
      recognizeBtn: document.getElementById("recognizeBtn"),
      clearBtn: document.getElementById("clearBtn"),
      previewBox: document.getElementById("previewBox"),
      comparePreviewBox: document.getElementById("comparePreviewBox"),
      compareStage: document.getElementById("compareStage"),
      fileBox: document.getElementById("fileBox"),
      tableSection: document.getElementById("tableSection"),
      outputBox: document.getElementById("outputBox"),
      main: document.querySelector(".main"),
      outputSection: document.querySelector(".output-section")
    };

    let rightPaneLayoutRaf = 0;

    function setOutput(text) {
      els.outputBox.textContent = text || "";
      els.outputBox.scrollTop = els.outputBox.scrollHeight;
    }

    function scheduleRightPaneLayoutSync() {
      if (rightPaneLayoutRaf) cancelAnimationFrame(rightPaneLayoutRaf);
      rightPaneLayoutRaf = requestAnimationFrame(() => {
        rightPaneLayoutRaf = requestAnimationFrame(() => {
          rightPaneLayoutRaf = 0;
        });
      });
    }

    function setModelSelectOpen(open) {
      if (!els.modelSelectRoot || !els.modelSelectTrigger) return;
      els.modelSelectRoot.classList.toggle("open", open);
      els.modelSelectTrigger.setAttribute("aria-expanded", open ? "true" : "false");
    }

    function closeModelSelect() {
      setModelSelectOpen(false);
    }

    function setModelSelectLabel(text) {
      if (els.modelSelectLabel) {
        els.modelSelectLabel.textContent = text || "未加载到模型";
      }
    }

    function setSelectedModel(modelId, label = modelId, { persist = true } = {}) {
      const next = String(modelId || "").trim();
      state.model = next;
      if (persist && next) {
        state.selectedModelByProvider[state.provider] = next;
      }
      if (els.modelSelect) {
        els.modelSelect.value = next;
      }
      setModelSelectLabel(label || (next ? next : "未加载到模型"));
      if (els.modelSelectMenu) {
        const buttons = els.modelSelectMenu.querySelectorAll(".model-select-option");
        buttons.forEach((button) => {
          button.classList.toggle("active", button.dataset.value === next);
        });
      }
      if (persist) persistConfig();
    }

    function formatTokensPerMillion(value) {
      const numeric = Number(value);
      if (!Number.isFinite(numeric)) return "未知";
      return `$${(numeric * 1000000).toFixed(numeric >= 0.01 ? 0 : 2)}`;
    }

    function getOpenRouterPriceLabel(model) {
      const pricing = model?.pricing || {};
      const prompt = formatTokensPerMillion(pricing.prompt);
      const completion = formatTokensPerMillion(pricing.completion);
      return `in ${prompt} / 1M · out ${completion} / 1M`;
    }

    function isOpenRouterAutoModel(model) {
      const id = String(model?.id || "").toLowerCase();
      const name = String(model?.name || "").toLowerCase();
      return id.startsWith("openrouter/auto") || name.includes("auto router");
    }

    function isOpenRouterFreeModel(model) {
      const pricing = model?.pricing || {};
      const priceFields = ["prompt", "completion", "request", "image", "web_search", "internal_reasoning"];
      const hasPaidPrice = priceFields.some((field) => Number(pricing[field] || 0) > 0);
      const name = String(model?.name || "").toLowerCase();
      const id = String(model?.id || "").toLowerCase();
      return !hasPaidPrice || name.includes("(free)") || id.includes(":free");
    }

    function isOpenRouterVisionModel(model) {
      const inputs = model?.architecture?.input_modalities || [];
      const outputs = model?.architecture?.output_modalities || [];
      return inputs.includes("image") && outputs.includes("text");
    }

    function normalizeOpenRouterModel(model) {
      const pricing = model?.pricing || {};
      const sortPrice = Number(pricing.prompt || 0) + Number(pricing.completion || 0) + Number(pricing.request || 0);
      return {
        id: String(model?.id || "").trim(),
        label: String(model?.name || model?.id || "").trim(),
        meta: getOpenRouterPriceLabel(model),
        sortPrice
      };
    }

    function getOpenRouterProviderRank(model) {
      const id = String(model?.id || "").toLowerCase();
      const label = String(model?.label || model?.name || "").toLowerCase();
      if (id.includes("qwen") || label.includes("qwen")) return 0;
      return 1;
    }

    function renderModelOptions(models, { emptyLabel = "未加载到模型" } = {}) {
      if (!els.modelSelectMenu) return;
      if (!models.length) {
        els.modelSelectMenu.innerHTML = `<div class="model-select-empty">${escapeHtml(emptyLabel)}</div>`;
        setSelectedModel("", emptyLabel, { persist: false });
        return;
      }

      els.modelSelectMenu.innerHTML = models.map((model) => {
        const active = model.id === state.model ? " active" : "";
        return `
          <button type="button" class="model-select-option${active}" role="option" data-value="${escapeHtml(model.id)}" data-label="${escapeHtml(model.label)}">
            <span class="model-select-option-inner">
              <span class="model-select-option-name">${escapeHtml(model.label)}</span>
              <span class="model-select-option-meta">${escapeHtml(model.meta || "")}</span>
            </span>
          </button>`;
      }).join("");

      const visibleModel = models.find((item) => item.id === state.model) || models[0];
      setSelectedModel(visibleModel.id, visibleModel.label, { persist: false });
    }

    function isEmbeddingsModel(name) {
      return /embed|embedding/i.test(String(name || ""));
    }

    function updateZoomIndicator() {
      // Zoom indicator removed from the UI; keep this as a no-op for existing flow.
    }

    function clampCompareOffsets() {
      const img = els.comparePreviewBox.querySelector("img");
      if (!img) return;
      const stageWidth = els.compareStage.clientWidth || 1;
      const stageHeight = els.compareStage.clientHeight || 1;
      const renderedWidth = (img.naturalWidth || 0) * state.compareView.fitScale * state.compareView.zoom;
      const renderedHeight = (img.naturalHeight || 0) * state.compareView.fitScale * state.compareView.zoom;
      const maxOffsetX = Math.max(0, (renderedWidth - stageWidth) / 2);
      const maxOffsetY = Math.max(0, (renderedHeight - stageHeight) / 2);
      state.compareView.offsetX = Math.min(maxOffsetX, Math.max(-maxOffsetX, state.compareView.offsetX));
      state.compareView.offsetY = Math.min(maxOffsetY, Math.max(-maxOffsetY, state.compareView.offsetY));
    }

    function refreshCompareFit({ resetOffsets = false } = {}) {
      const img = els.comparePreviewBox.querySelector("img");
      if (!img || !img.naturalWidth || !img.naturalHeight) {
        state.compareView.fitScale = 1;
        updateZoomIndicator();
        return;
      }
      const stageWidth = els.compareStage.clientWidth || 1;
      const stageHeight = els.compareStage.clientHeight || 1;
      const scaleX = stageWidth / img.naturalWidth;
      const scaleY = stageHeight / img.naturalHeight;
      state.compareView.fitScale = Math.min(scaleX, scaleY);
      if (resetOffsets || state.compareView.zoom <= 1) {
        state.compareView.offsetX = 0;
        state.compareView.offsetY = 0;
      } else {
        clampCompareOffsets();
      }
      applyCompareTransform();
    }

    function applyCompareTransform() {
      const img = els.comparePreviewBox.querySelector("img");
      updateZoomIndicator();
      if (!img) return;
      clampCompareOffsets();
      const displayScale = state.compareView.fitScale * state.compareView.zoom;
      img.style.transform = `translate3d(${state.compareView.offsetX}px, ${state.compareView.offsetY}px, 0) translate(-50%, -50%) scale(${displayScale})`;
      img.style.transformOrigin = "center center";
    }

    function resetCompareTransform(zoom = BASE_COMPARE_ZOOM) {
      state.compareView.zoom = zoom;
      state.compareView.offsetX = 0;
      state.compareView.offsetY = 0;
      refreshCompareFit({ resetOffsets: true });
    }

    function triggerPreviewReveal() {
      if (!els.previewBox) return;
      els.previewBox.classList.remove("is-revealing");
      void els.previewBox.offsetWidth;
      els.previewBox.classList.add("is-revealing");
      els.previewBox.addEventListener("animationend", () => {
        els.previewBox.classList.remove("is-revealing");
      }, { once: true });
    }

    function setCompareScale(nextZoom) {
      const zoom = Math.max(0.5, Math.min(6, nextZoom));
      state.compareView.zoom = zoom;
      if (zoom <= 1) {
        state.compareView.offsetX = 0;
        state.compareView.offsetY = 0;
      }
      applyCompareTransform();
    }

    function syncImagePreviewPanels() {
      const emptyHtml = '<span class="muted">点击或拖拽图片到此处</span>';
      const compareEmptyHtml = '<span class="compare-empty">等待上传图片。</span>';
      if (state.imageDataUrl) {
        const imageHtml = `<img src="${state.imageDataUrl}" alt="预览" draggable="false">`;
        els.previewBox.innerHTML = imageHtml;
        triggerPreviewReveal();
        els.comparePreviewBox.innerHTML = imageHtml;
        els.comparePreviewBox.classList.add("has-image");
        els.comparePreviewBox.classList.remove("is-empty");
        els.fileBox.classList.add("has-image");
        const compareImg = els.comparePreviewBox.querySelector("img");
        if (compareImg) {
          compareImg.addEventListener("load", () => refreshCompareFit({ resetOffsets: true }), { once: true });
          if (compareImg.complete) refreshCompareFit({ resetOffsets: true });
        }
        resetCompareTransform();
      } else {
        els.previewBox.innerHTML = emptyHtml;
        els.previewBox.classList.remove("is-revealing");
        els.comparePreviewBox.innerHTML = compareEmptyHtml;
        els.comparePreviewBox.classList.remove("has-image");
        els.comparePreviewBox.classList.add("is-empty");
        els.fileBox.classList.remove("has-image");
        resetCompareTransform();
      }
    }

    function bindCompareViewerEvents() {
      els.compareStage.addEventListener("wheel", (event) => {
        if (!state.imageDataUrl) return;
        event.preventDefault();
        const ratio = event.deltaY < 0 ? 1.12 : 1 / 1.12;
        setCompareScale(state.compareView.zoom * ratio);
      }, { passive: false });

      els.compareStage.addEventListener("pointerdown", (event) => {
        if (!state.imageDataUrl) return;
        if (event.button !== 0) return;
        if (state.compareView.zoom <= 1) return;
        event.preventDefault();
        state.compareView.dragging = true;
        state.compareView.startX = event.clientX;
        state.compareView.startY = event.clientY;
        state.compareView.baseX = state.compareView.offsetX;
        state.compareView.baseY = state.compareView.offsetY;
        els.compareStage.classList.add("is-dragging");
        els.compareStage.setPointerCapture(event.pointerId);
      });

      els.compareStage.addEventListener("pointermove", (event) => {
        if (!state.compareView.dragging) return;
        event.preventDefault();
        state.compareView.offsetX = state.compareView.baseX + (event.clientX - state.compareView.startX);
        state.compareView.offsetY = state.compareView.baseY + (event.clientY - state.compareView.startY);
        applyCompareTransform();
      });

      const stopDragging = (event) => {
        if (state.compareView.dragging && event?.pointerId != null && els.compareStage.hasPointerCapture(event.pointerId)) {
          els.compareStage.releasePointerCapture(event.pointerId);
        }
        state.compareView.dragging = false;
        els.compareStage.classList.remove("is-dragging");
      };

      els.compareStage.addEventListener("pointerup", stopDragging);
      els.compareStage.addEventListener("pointercancel", stopDragging);
      els.compareStage.addEventListener("lostpointercapture", stopDragging);
      els.compareStage.addEventListener("dblclick", () => {
        if (!state.imageDataUrl) return;
        resetCompareTransform();
      });
    }

    function setConnectionStatus(status) {
      const badge = els.connectionStatusBtn;
      if (!badge) return;
      badge.className = "btn-status";
      if (status === "connected") {
        badge.classList.add("connected");
        badge.textContent = "已连接";
        state.connected = true;
      } else if (status === "error") {
        badge.classList.add("error");
        badge.textContent = "连接失败";
        state.connected = false;
      } else {
        badge.textContent = "未检测";
        state.connected = false;
      }
    }

    function setRecognizing(busy) {
      els.recognizeBtn.disabled = busy;
      if (busy) {
        els.recognizeBtn.innerHTML = '<span class="spinner" style="display:inline-block;vertical-align:middle;"></span> 识别中...';
        els.recognizeBtn.classList.add("btn-processing");
      } else {
        els.recognizeBtn.textContent = "开始识别";
        els.recognizeBtn.classList.remove("btn-processing");
      }
    }

    function persistConfig() {
      localStorage.setItem(STORAGE_KEYS.baseUrl, state.baseUrl);
      localStorage.setItem(STORAGE_KEYS.model, state.model);
      localStorage.setItem(STORAGE_KEYS.provider, state.provider);
      localStorage.setItem(STORAGE_KEYS.openRouterApiKey, state.openRouterApiKey);
      localStorage.setItem(STORAGE_KEYS.selectedModelByProvider, JSON.stringify(state.selectedModelByProvider || {}));
    }

    function buildRecognitionSnapshot() {
      const imageDataUrl = state.imageDataUrl && state.imageDataUrl.length <= 1500000 ? state.imageDataUrl : "";
      return {
        version: 1,
        imageDataUrl,
        parsedRecords: state.parsedRecords,
        rawText: state.rawText,
        savedAt: Date.now()
      };
    }

    function persistRecognitionSnapshot() {
      try {
        if (!state.parsedRecords.length) {
          localStorage.removeItem(STORAGE_KEYS.recognitionSnapshot);
          return;
        }
        localStorage.setItem(STORAGE_KEYS.recognitionSnapshot, JSON.stringify(buildRecognitionSnapshot()));
      } catch (error) {
        console.warn("保存识别结果快照失败", error);
      }
    }

    function restoreRecognitionSnapshot() {
      const snapshot = readJsonStorage(STORAGE_KEYS.recognitionSnapshot, null);
      if (!snapshot || !Array.isArray(snapshot.parsedRecords) || snapshot.parsedRecords.length === 0) {
        return false;
      }
      state.imageDataUrl = typeof snapshot.imageDataUrl === "string" ? snapshot.imageDataUrl : "";
      state.rawText = typeof snapshot.rawText === "string" ? snapshot.rawText : "";
      state.parsedRecords = snapshot.parsedRecords;
      syncImagePreviewPanels();
      renderTable(state.parsedRecords);
      if (state.rawText) {
        setOutput(`已恢复 ${state.parsedRecords.length} 条记录\n\n` + state.rawText);
      } else {
        setOutput(`已恢复 ${state.parsedRecords.length} 条记录。`);
      }
      refreshCopyButtonState();
      return true;
    }

    function clearRecognitionSnapshot() {
      try {
        localStorage.removeItem(STORAGE_KEYS.recognitionSnapshot);
      } catch (error) {
        console.warn("清理识别结果快照失败", error);
      }
    }

    function normalizeBaseUrl(url) {
      return String(url || "").trim().replace(/\/+$/, "");
    }

    function normalizeProvider(provider) {
      return provider === "openrouter" ? "openrouter" : "lmstudio";
    }

    function setProvider(provider, { persist = true } = {}) {
      const next = normalizeProvider(provider);
      state.provider = next;
      if (els.providerSelect) {
        els.providerSelect.value = next;
      }
      if (persist) persistConfig();
    }

    function setOpenRouterApiKey(apiKey, { persist = true } = {}) {
      state.openRouterApiKey = String(apiKey || "").trim();
      if (els.openRouterKeyInput) {
        els.openRouterKeyInput.value = state.openRouterApiKey;
      }
      if (persist) persistConfig();
    }

    function getOpenRouterHeaders() {
      const headers = { Accept: "application/json" };
      if (state.openRouterApiKey) {
        headers.Authorization = `Bearer ${state.openRouterApiKey}`;
      }
      return headers;
    }

    function escapeHtml(value) {
      return String(value)
        .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;").replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
    }

    function normalizeFieldToken(value) {
      return String(value || "").toLowerCase()
        .replace(/[\s\r\n\t]/g, "").replace(/[℃【\[]/g, "[").replace(/[°】\]]/g, "]")
        .replace(/℃|°c/g, "c").replace(/[%％]/g, "%")
        .replace(/[:"'`.,，。？！\\|_-]/g, "")
        .replace(/\[(mpa|%|c)\]/g, "$1").replace(/\[|\]/g, "");
    }

    const FIELD_LOOKUP = (() => {
      const map = new Map();
      const register = (source, canonical) => {
        map.set(source, canonical);
        map.set(normalizeFieldToken(source), canonical);
      };
      COLUMN_DEFS.forEach((column) => register(column.key, column.key));
      Object.entries(FIELD_ALIASES).forEach(([alias, canonical]) => register(alias, canonical));
      return map;
    })();

    const CANONICAL_FIELD_TOKENS = COLUMN_DEFS.map((column) => ({
      canonical: column.key,
      token: normalizeFieldToken(column.key)
    }));

    function resolveCanonicalFieldName(rawKey) {
      if (FIELD_LOOKUP.has(rawKey)) return FIELD_LOOKUP.get(rawKey);
      const token = normalizeFieldToken(rawKey);
      if (FIELD_LOOKUP.has(token)) return FIELD_LOOKUP.get(token);
      const fuzzy = CANONICAL_FIELD_TOKENS.find((item) => token.includes(item.token) || item.token.includes(token));
      return fuzzy ? fuzzy.canonical : rawKey;
    }

    function mergeFieldValue(leftValue, rightValue) {
      if (leftValue == null || leftValue === "") return rightValue;
      if (rightValue == null || rightValue === "") return leftValue;
      if (Array.isArray(leftValue) || Array.isArray(rightValue)) {
        const left = Array.isArray(leftValue) ? leftValue : [leftValue];
        const right = Array.isArray(rightValue) ? rightValue : [rightValue];
        return [...left, ...right].filter((item) => item !== "" && item != null);
      }
      return leftValue;
    }

    function normalizeRecord(record) {
      const normalized = {};
      Object.entries(record || {}).forEach(([key, value]) => {
        const canonicalKey = resolveCanonicalFieldName(key);
        normalized[canonicalKey] = mergeFieldValue(normalized[canonicalKey], value);
      });
      if (Object.prototype.hasOwnProperty.call(normalized, "型号")) {
        normalized["型号"] = normalizeModelValue(normalized["型号"]);
      }
      return normalized;
    }

    function normalizeModelValue(value) {
      const text = String(value ?? "").trim();
      if (!text) return "";
      return text
        .replace(/^(310|320|420|520)6/, "$1G")
        .replace(/-13/g, "-B");
    }

    function normalizeRecordTypes(record) {
      const template = REFERENCE_DATA[0];
      const output = { ...record };
      Object.keys(template).forEach((key) => {
        const shouldBeArray = Array.isArray(template[key]);
        const value = output[key];
        if (value == null || value === "") { output[key] = shouldBeArray ? [] : ""; return; }
        if (shouldBeArray && !Array.isArray(value)) { output[key] = [value]; return; }
        if (!shouldBeArray && Array.isArray(value)) { output[key] = String(value[0] ?? ""); }
      });
      return output;
    }

    function hasMeaningfulValue(value) {
      if (value == null) return false;
      if (Array.isArray(value)) return value.some((item) => hasMeaningfulValue(item));
      if (typeof value === "object") return Object.values(value).some((item) => hasMeaningfulValue(item));
      return String(value).trim() !== "";
    }

    function isMeaningfulRecord(record) {
      if (!record || typeof record !== "object") return false;
      return Object.entries(record).some(([key, value]) => {
        if (key === "型号") return hasMeaningfulValue(value);
        return hasMeaningfulValue(value);
      });
    }

    function extractModelCandidates(rawText) {
      const text = String(rawText || "");
      const matches = new Set();
      const addMatch = (value) => {
        const normalized = normalizeModelValue(value);
        if (normalized) matches.add(normalized);
      };

      const pattern = /(?:310|320|420|520)\s*[0-9A-Za-z\-]{2,}/g;
      for (const match of text.matchAll(pattern)) {
        addMatch(match[0]);
      }

      const quotedPattern = /"([^"]+)"/g;
      for (const match of text.matchAll(quotedPattern)) {
        const content = match[1] || "";
        const innerMatches = content.match(pattern) || [];
        innerMatches.forEach(addMatch);
      }

      const tokens = text
        .split(/[\s,，;；:：|/\\()\[\]{}<>]+/)
        .map((token) => token.trim())
        .filter(Boolean);
      tokens
        .filter((token) => /^(?:310|320|420|520)/.test(token))
        .forEach(addMatch);

      const candidates = [...matches];
      return [...new Set(candidates)];
    }

    function expandSingleRecordByModelCandidates(rawText, records) {
      if (!Array.isArray(records) || records.length !== 1) return records;
      const candidates = extractModelCandidates(rawText);
      if (candidates.length <= 1) return records;

      const source = records[0];
      const arrayKeys = Object.entries(source)
        .filter(([key, value]) => key !== "型号" && Array.isArray(value))
        .map(([key]) => key);
      const maxRows = Math.max(candidates.length, ...arrayKeys.map((key) => source[key].length));
      return candidates.map((model, index) => {
        const cloned = {};
        Object.entries(source).forEach(([key, value]) => {
          if (key === "型号") {
            cloned[key] = model;
            return;
          }
          if (Array.isArray(value)) {
            cloned[key] = value[index] != null ? value[index] : value[Math.min(value.length - 1, index)] ?? "";
            return;
          }
          cloned[key] = value;
        });
        cloned["型号"] = normalizeModelValue(cloned["型号"]);
        return normalizeRecordTypes(cloned);
      }).slice(0, maxRows);
    }

    function extractJsonCandidate(rawText) {
      const fenced = rawText.match(/```json\s*([\s\S]*?)```/i) || rawText.match(/```\s*([\s\S]*?)```/i);
      if (fenced) return fenced[1].trim();
      const firstBracket = rawText.indexOf("["), lastBracket = rawText.lastIndexOf("]");
      if (firstBracket !== -1 && lastBracket !== -1 && lastBracket > firstBracket) return rawText.slice(firstBracket, lastBracket + 1);
      const firstBrace = rawText.indexOf("{"), lastBrace = rawText.lastIndexOf("}");
      if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) return "[" + rawText.slice(firstBrace, lastBrace + 1) + "]";
      return rawText.trim();
    }

    function parseModelJson(rawText) {
      const candidate = extractJsonCandidate(rawText);
      let parsed = JSON.parse(candidate);
      if (!Array.isArray(parsed)) parsed = [parsed];
      const normalized = parsed
        .map((record) => normalizeRecord(record))
        .map((record) => normalizeRecordTypes(record))
        .filter((record) => isMeaningfulRecord(record));
      return expandSingleRecordByModelCandidates(rawText, normalized);
    }

    function toSeries(value) {
      if (Array.isArray(value)) return value.map((item) => item == null ? "" : String(item));
      if (value == null || value === "") return [];
      return [String(value)];
    }

    function formatCell(value) {
      const text = String(value ?? "").trim();
      return text ? text : "—";
    }

    function getEditorDisplayValue(value) {
      const text = String(value ?? "").trim();
      return text || "—";
    }

    function isTableCellEditor(target) {
      return target instanceof HTMLElement && target.classList.contains("table-cell-editor");
    }

    function setTableCellValue(recordIndex, key, value, arrayIndex = null) {
      const record = state.parsedRecords[recordIndex];
      if (!record) return;
      const nextValue = key === "型号" ? normalizeModelValue(value) : String(value ?? "").trim();

      if (arrayIndex == null) {
        record[key] = nextValue;
        persistRecognitionSnapshot();
        return;
      }

      if (!Array.isArray(record[key])) {
        record[key] = [];
      }
      const target = record[key];
      while (target.length <= arrayIndex) {
        target.push("");
      }
      target[arrayIndex] = nextValue;
      persistRecognitionSnapshot();
    }

    function syncEditedJsonOutput() {
      try {
        setOutput(JSON.stringify(state.parsedRecords, null, 2));
      } catch {
        setOutput("当前表格内容无法序列化为 JSON。");
      }
    }

    function buildTableExportMatrix(records) {
      const columns = getOrderedColumns(records);
      const seriesColumns = getSeriesColumns(records, columns);
      const scalarColumns = columns.filter((key) => !seriesColumns.includes(key));
      const impactColumnCount = getImpactColumnCount(records, seriesColumns);
      const visibleSeriesColumns = seriesColumns.flatMap((key) => {
        if (key !== "冲击强度[Mpa]") return [key];
        return Array.from({ length: impactColumnCount }, (_, index) => ({ key, partIndex: index }));
      });
      const headers = [
        ...scalarColumns,
        ...visibleSeriesColumns.map((column) => (typeof column === "string" ? column : column.key))
      ];

      const rows = [];
      records.forEach((record) => {
        const rowCount = getRowCount(record, seriesColumns);
        Array.from({ length: rowCount }, (_, rowIndex) => {
          const row = [];
          scalarColumns.forEach((key, scalarIndex) => {
            const mergeLeadingScalar = scalarIndex < 3 && rowIndex > 0;
            if (mergeLeadingScalar) {
              row.push("");
              return;
            }
            row.push(formatCell(record[key]) === "—" ? "" : String(record[key] ?? ""));
          });
          visibleSeriesColumns.forEach((column) => {
            const key = typeof column === "string" ? column : column.key;
            const series = toSeries(record[key]);
            let rawValue = "";
            if (typeof column === "string") {
              rawValue = series[rowIndex] ?? "";
            } else {
              const startIdx = column.partIndex * rowCount;
              rawValue = series[startIdx + rowIndex] ?? "";
            }
            row.push(String(rawValue ?? ""));
          });
          rows.push(row);
        });
      });

      return { headers, rows };
    }

    function buildClipboardPlainText(records) {
      const { rows } = buildTableExportMatrix(records);
      const sanitize = (value) => String(value ?? "").replace(/\r?\n/g, " ").replace(/\t/g, " ");
      return rows
        .map((row) => row.map(sanitize).join("\t"))
        .join("\r\n");
    }

    async function writePlainTextToClipboard(text) {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
        return;
      }
      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      document.execCommand("copy");
      textarea.remove();
    }

    async function copyRecordsToClipboard(records, label = "表格") {
      if (!records.length) {
        setOutput("没有可复制的表格内容。");
        return;
      }
      const activeEditor = document.activeElement;
      if (isTableCellEditor(activeEditor)) {
        commitTableCellEditor(activeEditor);
        activeEditor.blur();
      }
      const plainText = buildClipboardPlainText(records);
      try {
        if (navigator.clipboard && window.ClipboardItem) {
          const item = new ClipboardItem({
            "text/plain": new Blob([plainText], { type: "text/plain" })
          });
          await navigator.clipboard.write([item]);
        } else if (navigator.clipboard?.writeText) {
          await writePlainTextToClipboard(plainText);
        } else {
          await writePlainTextToClipboard(plainText);
        }
        setOutput(`${label}已复制到剪贴板。`);
      } catch (error) {
        setOutput("复制失败:\n\n" + (error.message || String(error)));
      }
    }

    async function copyRecordAtIndex(recordIndex) {
      const record = state.parsedRecords[recordIndex];
      if (!record) {
        setOutput("未找到要复制的数据组。");
        return;
      }
      const modelLabel = String(record["型号"] || "").trim() || `第 ${recordIndex + 1} 组`;
      await copyRecordsToClipboard([record], `${modelLabel} `);
    }

    function commitTableCellEditor(editor) {
      const recordIndex = Number(editor.dataset.recordIndex);
      const key = editor.dataset.key || "";
      const arrayIndex = editor.dataset.arrayIndex != null && editor.dataset.arrayIndex !== ""
        ? Number(editor.dataset.arrayIndex)
        : null;
      const rawValue = editor.textContent || "";
      const normalized = rawValue.trim() === "—" ? "" : rawValue.trim();
      const nextValue = key === "型号" ? normalizeModelValue(normalized) : normalized;
      editor.textContent = nextValue ? nextValue : "—";
      editor.classList.toggle("is-empty", !nextValue);
      if (Number.isInteger(recordIndex) && key) {
        setTableCellValue(recordIndex, key, nextValue, Number.isInteger(arrayIndex) ? arrayIndex : null);
        syncEditedJsonOutput();
      }
    }

    function clearEditorPlaceholder(editor) {
      if (!isTableCellEditor(editor)) return;
      if (editor.textContent.trim() === "—") {
        editor.textContent = "";
        editor.classList.remove("is-empty");
      }
    }

    function normalizeEditorPaste(event) {
      if (!isTableCellEditor(event.target)) return;
      event.preventDefault();
      const text = event.clipboardData?.getData("text/plain") || "";
      document.execCommand("insertText", false, text);
    }

    function getColumnClassName(key) {
      return COLUMN_DEFS.find((item) => item.key === key)?.className || "group-extra";
    }

    function getOrderedColumns(records) {
      const preferred = COLUMN_DEFS.map((column) => column.key);
      const extras = [], seen = new Set();
      records.forEach((record) => {
        Object.keys(record || {}).forEach((key) => {
          if (!seen.has(key)) { seen.add(key); if (!preferred.includes(key)) extras.push(key); }
        });
      });
      return [...preferred.filter((key) => seen.has(key)), ...extras];
    }

    function getSeriesColumns(records, columns) {
      return columns.filter((key) => records.some((record) => Array.isArray(record[key])));
    }

    function getPrimaryRowCount(record, seriesColumns) {
      const primarySeriesColumns = seriesColumns.filter((key) => key !== "冲击强度[Mpa]");
      const primaryMax = Math.max(0, ...primarySeriesColumns.map((key) => toSeries(record[key]).length));
      if (primaryMax > 0) return primaryMax;
      const impactCount = toSeries(record["冲击强度[Mpa]"]).length;
      if (impactCount > 0) return impactCount;
      return 1;
    }

    function getRowCount(record, seriesColumns) {
      return getPrimaryRowCount(record, seriesColumns);
    }

    function getImpactColumnCount(records, baseColumns) {
      if (!baseColumns.includes("冲击强度[Mpa]")) return 0;
      return Math.max(1, ...records.map((record) => {
        const rowCount = getPrimaryRowCount(record, baseColumns);
        const impactCount = toSeries(record["冲击强度[Mpa]"]).length;
        return Math.max(1, Math.ceil(impactCount / rowCount));
      }));
    }

    function renderEmptyTable(message) {
      els.tableSection.innerHTML = `
        <div class="empty-panel">
          <div class="empty-icon">📋</div>
          <div class="empty-text">${escapeHtml(message)}</div>
        </div>`;
      refreshCopyButtonState();
      scheduleRightPaneLayoutSync();
    }

    function syncRightPaneLayout() {
      const main = els.main, tableSection = els.tableSection, outputSection = els.outputSection;
      if (!main || !tableSection || !outputSection) return;
      tableSection.style.flex = "0 0 auto";
      tableSection.style.minHeight = "0";
      tableSection.style.height = "auto";
      outputSection.style.flex = "1 1 auto";
      outputSection.style.minHeight = "0";
      outputSection.style.height = "auto";
    }

    function renderTable(records) {
      if (!records.length) { renderEmptyTable("识别成功后，这里会按专业表格格式显示 JSON 数据。"); return; }
      const columns = getOrderedColumns(records);
      const seriesColumns = getSeriesColumns(records, columns);
      const scalarColumns = columns.filter((key) => !seriesColumns.includes(key));
      const impactColumnCount = getImpactColumnCount(records, seriesColumns);
      const visibleSeriesColumns = seriesColumns.flatMap((key) => {
        if (key !== "冲击强度[Mpa]") return [key];
        return Array.from({ length: impactColumnCount }, (_, index) => ({ key, partIndex: index }));
      });

      const header = [
        ...scalarColumns.map((key) => `<th>${escapeHtml(key)}</th>`),
        ...visibleSeriesColumns.map((column) => {
          if (typeof column === "string") return `<th>${escapeHtml(column)}</th>`;
          return `<th>${escapeHtml(column.key)}</th>`;
        })
      ].join("");

      const toolbar = records.map((record, recordIndex) => {
        const modelLabel = String(record["型号"] || "").trim() || `第 ${recordIndex + 1} 组`;
        const batchLabel = String(record["批次"] || "").trim();
        const extraLabel = batchLabel ? ` / ${batchLabel}` : "";
        return `<button type="button" class="btn-secondary table-copy-btn" data-copy-record-index="${recordIndex}">复制 ${escapeHtml(modelLabel + extraLabel)}</button>`;
      }).join("");
      const subtitle = `点按按钮复制单组数据，表格内容仍可直接编辑`;
      const countLabel = `共 ${records.length} 组`;

      const body = records.map((record, recordIndex) => {
        const rowCount = getRowCount(record, seriesColumns);
        return Array.from({ length: rowCount }, (_, rowIndex) => {
          const cells = [];
          const rowClass = rowIndex === 0 && recordIndex > 0 ? "record-start" : "";
          scalarColumns.forEach((key) => {
            if (rowIndex > 0) return;
            const value = formatCell(record[key]);
            const emptyClass = value === "—" ? " empty-cell" : "";
            cells.push(`<td class="${getColumnClassName(key)}" rowspan="${rowCount}">
              <span
                class="table-cell-editor ${emptyClass.trim() ? "is-empty" : ""}"
                contenteditable="true"
                spellcheck="false"
                role="textbox"
                aria-label="${escapeHtml(key)}"
                data-record-index="${recordIndex}"
                data-key="${escapeHtml(key)}">${escapeHtml(getEditorDisplayValue(record[key]))}</span>
            </td>`);
          });
          visibleSeriesColumns.forEach((column) => {
            const key = typeof column === "string" ? column : column.key;
            const className = getColumnClassName(key);
            const series = toSeries(record[key]);
            let rawValue = "";
            if (typeof column === "string") {
              rawValue = series[rowIndex] ?? "";
            } else {
              const startIdx = column.partIndex * rowCount;
              rawValue = series[startIdx + rowIndex] ?? "";
            }
            const value = formatCell(rawValue);
            const emptyClass = value === "—" ? " empty-cell" : "";
            const arrayIndex = typeof column === "string" ? rowIndex : column.partIndex * rowCount + rowIndex;
            cells.push(`<td class="${className}">
              <span
                class="table-cell-editor ${emptyClass.trim() ? "is-empty" : ""}"
                contenteditable="true"
                spellcheck="false"
                role="textbox"
                aria-label="${escapeHtml(key)}"
                data-record-index="${recordIndex}"
                data-key="${escapeHtml(key)}"
                data-array-index="${arrayIndex}">${escapeHtml(getEditorDisplayValue(rawValue))}</span>
            </td>`);
          });
          return `<tr class="${rowClass}">${cells.join("")}</tr>`;
        }).join("");
      }).join("");

      els.tableSection.innerHTML = `
        <div class="table-meta">
          <div class="table-meta-copy">
            <div class="table-title-row">
              <div class="table-title">识别结果</div>
              <div class="table-count-pill">${escapeHtml(countLabel)}</div>
            </div>
            <div class="table-subtitle">${escapeHtml(subtitle)}</div>
          </div>
          <div class="table-toolbar">${toolbar}</div>
        </div>
        <div class="table-wrap">
          <table class="report-table">
            <thead><tr>${header}</tr></thead>
            <tbody>${body}</tbody>
          </table>
        </div>`;
      refreshCopyButtonState();
      scheduleRightPaneLayoutSync();
    }

    function fileToDataUrl(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error || new Error("图片读取失败"));
        reader.readAsDataURL(file);
      });
    }

    async function loadLocalModels({ silent = false } = {}) {
      state.baseUrl = els.baseUrlInput.value.trim();
      persistConfig();
      if (els.modelSelectRoot) els.modelSelectRoot.classList.add("is-loading");
      try {
        const response = await fetch(normalizeBaseUrl(state.baseUrl) + "/v1/models");
        if (!response.ok) throw new Error(`状态码 ${response.status}`);
        const payload = await response.json();
        const models = Array.isArray(payload.data)
          ? payload.data
              .map((item) => item?.id || item?.name || item?.model || "")
              .filter((name) => Boolean(name) && !isEmbeddingsModel(name))
              .map((name) => ({ id: name, label: name, meta: "LM Studio" }))
          : [];
        if (els.modelSelectRoot) els.modelSelectRoot.classList.remove("is-loading");
        if (!models.length) {
          renderModelOptions([], { emptyLabel: "未加载到模型" });
          setConnectionStatus("connected");
          return;
        }
        renderModelOptions(models, { emptyLabel: "未加载到模型" });
        const remembered = state.selectedModelByProvider.lmstudio || state.model;
        const next = models.find((item) => item.id === remembered) || models[0];
        setSelectedModel(next.id, next.label);
        setConnectionStatus("connected");
      } catch (error) {
        if (els.modelSelectRoot) els.modelSelectRoot.classList.remove("is-loading");
        renderModelOptions([], { emptyLabel: "未加载到模型" });
        setConnectionStatus("error");
        if (!silent) setOutput("加载模型失败:\n\n" + (error.message || String(error)));
      }
    }

    async function loadOpenRouterModels({ silent = false } = {}) {
      if (els.modelSelectRoot) els.modelSelectRoot.classList.add("is-loading");
      try {
        const response = await fetch("https://openrouter.ai/api/v1/models?output_modalities=text", {
          headers: getOpenRouterHeaders()
        });
        if (!response.ok) throw new Error(`状态码 ${response.status}`);
        const payload = await response.json();
        const models = Array.isArray(payload.data)
          ? payload.data
              .filter((item) => isOpenRouterVisionModel(item))
              .filter((item) => !isOpenRouterAutoModel(item))
              .filter((item) => !isOpenRouterFreeModel(item))
              .map(normalizeOpenRouterModel)
          : [];
        if (els.modelSelectRoot) els.modelSelectRoot.classList.remove("is-loading");
        if (!models.length) {
          renderModelOptions([], { emptyLabel: "未找到可分析图片的 OpenRouter 模型" });
          setConnectionStatus("connected");
          return;
        }
        models.sort((left, right) => {
          const leftRank = getOpenRouterProviderRank(left);
          const rightRank = getOpenRouterProviderRank(right);
          if (leftRank !== rightRank) return leftRank - rightRank;
          if (left.sortPrice !== right.sortPrice) return left.sortPrice - right.sortPrice;
          return left.label.localeCompare(right.label, "zh-Hans-CN");
        });
        renderModelOptions(models, { emptyLabel: "未找到可分析图片的 OpenRouter 模型" });
        const remembered = state.selectedModelByProvider.openrouter || state.model;
        const next = models.find((item) => item.id === remembered) || models[0];
        setSelectedModel(next.id, next.label);
        setConnectionStatus("connected");
      } catch (error) {
        if (els.modelSelectRoot) els.modelSelectRoot.classList.remove("is-loading");
        renderModelOptions([], { emptyLabel: "未找到可分析图片的 OpenRouter 模型" });
        setConnectionStatus("error");
        if (!silent) setOutput("加载 OpenRouter 模型失败:\n\n" + (error.message || String(error)));
      }
    }

    async function loadModels({ silent = false } = {}) {
      if (state.provider === "openrouter") {
        await loadOpenRouterModels({ silent });
        return;
      }
      await loadLocalModels({ silent });
    }

    function buildImageRecognitionMessages(imageDataUrl) {
      return [
        {
          role: "system",
          content: [
            "只输出格式化 JSON 数组，2 空格缩进，禁止 markdown、解释和代码块。",
            "识别到几组数据就返回几组数据；某一组如果没有任何有效字段、字段全为空、全是空数组、null 或无法确认的内容，就不要输出这一组。",
            "只保留有实际识别结果的组，不要补空组，也不要为了凑格式输出空对象。",
            "严格保留原始字符，不要把字母 B 误识别成 13，不要把字母 G 误识别成 6。",
            "测试温度字段只能输出 250°、260°、275° 这三种格式之一，必须是数字加中文角度符号，不要写成 250、250C、250℃ 或其它变体。",
            "如果一张图片里有多个不同型号的数据块，每个数据块必须单独输出为一个 JSON 对象，不要把多个型号合并进同一个对象。",
            "如果温度看不清，优先根据同一张表的上下文判断；实在无法判断就留空，不要瞎猜。"
          ].join(" ")
        },
        {
          role: "user",
          content: [
            { type: "text", text: "提取图片中的材料检测表，输出格式化 JSON 数组。识别到几组就返回几组；没有任何有效内容的组不要返回。每条记录只保留：型号、批次、测试温度、熔指、拉伸强度[Mpa]、断裂伸长率[%]、弯曲强度[Mpa]、弯曲模量[Mpa]、冲击强度[Mpa]。型号中的字母 B、G、I、S、O、Z 必须保持字母，不要和数字混淆；遇到模糊字符时不要强行替换。测试温度只能输出 250°、260°、275° 之一。基础字段输出字符串，测量字段输出数组，数组和对象都保持多行缩进。空字段可以省略，不要补空组或空对象。" },
            { type: "image_url", image_url: { url: imageDataUrl } }
          ]
        }
      ];
    }

    async function callStreamingChatCompletion(endpoint, headers, model, imageDataUrl, providerName, onDelta = () => {}) {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Accept": "text/event-stream", ...headers },
        body: JSON.stringify({
          model,
          temperature: 0.1,
          stream: true,
          messages: buildImageRecognitionMessages(imageDataUrl)
        })
      });

      if (!response.ok) throw new Error(`${providerName} 调用失败 (${response.status})\n${await response.text()}`);

      const reader = response.body?.getReader();
      if (!reader) {
        const payload = await response.json();
        const content = payload?.choices?.[0]?.message?.content || "";
        const text = Array.isArray(content) ? content.map((item) => item?.text || "").join("\n") : String(content);
        onDelta(text);
        return text;
      }

      const decoder = new TextDecoder("utf-8");
      let buffer = "", fullText = "";
      const flushChunk = (chunk) => {
        const lines = chunk.split(/\r?\n/);
        for (const line of lines) {
          if (!line.startsWith("data:")) continue;
          const data = line.slice(5).trim();
          if (!data || data === "[DONE]") continue;
          try {
            const payload = JSON.parse(data);
            const delta = payload?.choices?.[0]?.delta?.content
              ?? payload?.choices?.[0]?.message?.content
              ?? payload?.choices?.[0]?.text ?? "";
            if (delta) { fullText += delta; onDelta(fullText); }
          } catch { /* ignore */ }
        }
      };

      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const parts = buffer.split(/\r?\n\r?\n/);
        buffer = parts.pop() || "";
        parts.forEach(flushChunk);
      }
      buffer += decoder.decode();
      if (buffer.trim()) flushChunk(buffer);
      return fullText;
    }

    async function callLmStudio(baseUrl, model, imageDataUrl, onDelta = () => {}) {
      return callStreamingChatCompletion(
        normalizeBaseUrl(baseUrl) + "/v1/chat/completions",
        {},
        model,
        imageDataUrl,
        "LM Studio",
        onDelta
      );
    }

    async function callOpenRouter(model, imageDataUrl, onDelta = () => {}) {
      if (!state.openRouterApiKey) {
        throw new Error("请先填写 OpenRouter API Key");
      }
      return callStreamingChatCompletion(
        "https://openrouter.ai/api/v1/chat/completions",
        getOpenRouterHeaders(),
        model,
        imageDataUrl,
        "OpenRouter",
        onDelta
      );
    }

    async function recognizeImage() {
      state.baseUrl = els.baseUrlInput.value.trim();
      state.provider = normalizeProvider(els.providerSelect?.value || state.provider);
      state.openRouterApiKey = els.openRouterKeyInput?.value.trim() || state.openRouterApiKey;
      state.model = els.modelSelect.value.trim();
      persistConfig();
      if (!state.imageDataUrl) { setOutput("请先选择图片。"); return; }
      if (!state.model) { setOutput("请先选择模型。"); return; }

      setRecognizing(true);
      setOutput("开始识别...\n");
      try {
        const rawText = state.provider === "openrouter"
          ? await callOpenRouter(state.model, state.imageDataUrl, (text) => setOutput(text))
          : await callLmStudio(state.baseUrl, state.model, state.imageDataUrl, (text) => setOutput(text));
        state.rawText = rawText;
        state.parsedRecords = parseModelJson(rawText);
        renderTable(state.parsedRecords);
        persistRecognitionSnapshot();
        setOutput(`已解析 ${state.parsedRecords.length} 条记录\n\n` + (rawText || "模型没有返回内容。"));
      } catch (error) {
        state.parsedRecords = [];
        renderEmptyTable("识别失败，暂时无法生成表格。");
        clearRecognitionSnapshot();
        setOutput("识别失败:\n\n" + (error.message || String(error)));
      } finally {
        setRecognizing(false);
      }
    }

    function clearAll() {
      state.imageDataUrl = "";
      state.parsedRecords = [];
      state.rawText = "";
      els.imageInput.value = "";
      clearRecognitionSnapshot();
      syncImagePreviewPanels();
      renderEmptyTable("识别成功后，这里会按专业表格格式显示 JSON 数据。");
      setOutput("等待识别结果。");
      refreshCopyButtonState();
    }

    function refreshCopyButtonState() {
      const buttons = els.tableSection?.querySelectorAll("[data-copy-record-index]");
      buttons?.forEach((button) => {
        button.disabled = !state.parsedRecords.length;
      });
    }

    function bindEvents() {
      els.providerSelect.addEventListener("change", async (event) => {
        setProvider(event.target.value, { persist: true });
        await loadModels({ silent: false });
      });

      els.baseUrlInput.addEventListener("change", () => {
        state.baseUrl = els.baseUrlInput.value.trim();
        persistConfig();
      });

      els.openRouterKeyInput.addEventListener("input", () => {
        setOpenRouterApiKey(els.openRouterKeyInput.value, { persist: true });
      });

      els.checkConnectionBtn.addEventListener("click", () => loadModels({ silent: false }));
      els.recognizeBtn.addEventListener("click", recognizeImage);
      els.clearBtn.addEventListener("click", clearAll);
      bindCompareViewerEvents();
      els.tableSection.addEventListener("click", async (event) => {
        const button = event.target.closest("[data-copy-record-index]");
        if (!button) return;
        const recordIndex = Number(button.dataset.copyRecordIndex);
        if (!Number.isInteger(recordIndex)) return;
        await copyRecordAtIndex(recordIndex);
      });
      els.tableSection.addEventListener("focusin", (event) => {
        clearEditorPlaceholder(event.target);
      });
      els.tableSection.addEventListener("focusout", (event) => {
        if (isTableCellEditor(event.target)) {
          commitTableCellEditor(event.target);
        }
      });
      els.tableSection.addEventListener("keydown", (event) => {
        if (!isTableCellEditor(event.target)) return;
        if (event.key === "Enter") {
          event.preventDefault();
          event.target.blur();
        }
      });
      els.tableSection.addEventListener("paste", normalizeEditorPaste);
      els.modelSelectTrigger.addEventListener("click", (event) => {
        event.preventDefault();
        if (els.modelSelectRoot?.classList.contains("is-loading")) return;
        const isOpen = els.modelSelectRoot?.classList.contains("open");
        setModelSelectOpen(!isOpen);
      });

      els.modelSelectMenu.addEventListener("click", (event) => {
        const button = event.target.closest(".model-select-option");
        if (!button) return;
        setSelectedModel(button.dataset.value || "", button.dataset.label || button.dataset.value || "");
        closeModelSelect();
      });

      document.addEventListener("click", (event) => {
        if (!els.modelSelectRoot || els.modelSelectRoot.contains(event.target)) return;
        closeModelSelect();
      });

      document.addEventListener("keydown", (event) => {
        if (event.key === "Escape") closeModelSelect();
      });

      els.imageInput.addEventListener("change", async (event) => {
        const file = event.target.files?.[0];
        if (!file) return;
        try {
          state.imageDataUrl = await fileToDataUrl(file);
          syncImagePreviewPanels();
          setOutput("图片已加载。");
        } catch (error) {
          setOutput("图片读取失败:\n\n" + (error.message || String(error)));
        }
      });

      // Drag & drop support
      const fileBox = els.fileBox;
      fileBox.addEventListener("dragover", (e) => { e.preventDefault(); fileBox.style.borderColor = "var(--primary)"; });
      fileBox.addEventListener("dragleave", () => { fileBox.style.borderColor = ""; });
      fileBox.addEventListener("drop", async (e) => {
        e.preventDefault();
        fileBox.style.borderColor = "";
        const file = e.dataTransfer?.files?.[0];
        if (!file || !file.type.startsWith("image/")) return;
        try {
          state.imageDataUrl = await fileToDataUrl(file);
          syncImagePreviewPanels();
          setOutput("图片已加载（拖拽）。");
        } catch (error) {
          setOutput("图片读取失败:\n\n" + (error.message || String(error)));
        }
      });
    }

    function init() {
      els.baseUrlInput.value = state.baseUrl;
      setProvider(state.provider, { persist: false });
      setOpenRouterApiKey(state.openRouterApiKey, { persist: false });
      state.model = state.selectedModelByProvider[state.provider] || state.model || localStorage.getItem(STORAGE_KEYS.model) || "";
      setModelSelectLabel(state.model || "请先检测连接以加载模型列表");
      if (els.modelSelect) els.modelSelect.value = state.model || "";
      bindEvents();
      setSelectedModel(state.model || "", state.model || "请先检测连接以加载模型列表", { persist: false });
      loadModels({ silent: true });
      const restored = restoreRecognitionSnapshot();
      if (!restored) {
        syncImagePreviewPanels();
        renderEmptyTable("识别成功后，这里会按专业表格格式显示 JSON 数据。");
        setOutput("等待识别结果。");
      }
      refreshCopyButtonState();
      scheduleRightPaneLayoutSync();
    }

    init();
    window.addEventListener("load", () => {
      scheduleRightPaneLayoutSync();
      refreshCompareFit({ resetOffsets: state.compareView.zoom <= 1 });
    });
    window.addEventListener("resize", () => {
      scheduleRightPaneLayoutSync();
      refreshCompareFit({ resetOffsets: state.compareView.zoom <= 1 });
    });
  
