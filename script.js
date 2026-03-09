const STORAGE_KEY = "flowchart-builder-interactive:v1";
const HISTORY_STORAGE_KEY = "flowchart-builder-history:v1";
const DEFAULT_LAYOUT = "TB";
const PDF_WORKER_SRC =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const HISTORY_DATE_FORMATTER = new Intl.DateTimeFormat("en-US", {
  dateStyle: "medium",
  timeStyle: "short",
});
const SAMPLE_PROMPTS = {
  expenseApproval: {
    label: "Example 1 - Expense Approval",
    text: [
      "An employee submits an expense request.",
      "The finance team reviews the request.",
      "",
      "If the expense follows company policy, it is approved and payment is issued.",
      "If the expense does not follow policy, the request is returned to the employee for correction.",
      "",
      "Once corrected, the employee resubmits the request for review.",
    ].join("\n"),
  },
  supportTicket: {
    label: "Example 2 - Support Ticket Process",
    text: [
      "A user reports an issue through the support portal.",
      "The support team reviews the ticket.",
      "",
      "If the issue is simple, the support team resolves it and closes the ticket.",
      "If the issue is complex, the ticket is escalated to the technical team.",
      "",
      "Once the issue is fixed, the user is notified and the ticket is closed.",
    ].join("\n"),
  },
  contentPublishing: {
    label: "Example 3 - Content Publishing Workflow",
    text: [
      "A writer submits an article draft.",
      "An editor reviews the article.",
      "",
      "If the article needs changes, it is returned to the writer for revision.",
      "If the article is approved, it is scheduled for publishing.",
      "",
      "Once published, the article is shared on the website.",
    ].join("\n"),
  },
  employeeOnboarding: {
    label: "Example 4 - Employee Onboarding Workflow",
    text: [
      "A new employee submits onboarding documents.",
      "HR reviews the documents.",
      "",
      "If any documents are missing, HR returns them to the employee for correction.",
      "If the documents are complete, HR creates the employee record.",
      "",
      "Once the employee record is ready, IT sets up system access and equipment.",
      "If setup is delayed, the request remains in progress.",
      "",
      "When IT setup is complete, the manager assigns onboarding tasks.",
      "Once the tasks are assigned, the employee is fully onboarded.",
    ].join("\n"),
  },
};
const NODE_COLORS = {
  start: {
    fill: "#FFFFFF",
    stroke: "#12345B",
    shape: "ellipse",
  },
  process: {
    fill: "#FFFFFF",
    stroke: "#12345B",
    shape: "rectangle",
  },
  decision: {
    fill: "#FFFFFF",
    stroke: "#12345B",
    shape: "diamond",
  },
  end: {
    fill: "#FFFFFF",
    stroke: "#12345B",
    shape: "ellipse",
  },
};

const refs = {};
const state = {
  prompt: "",
  code: "",
  graph: null,
  layout: DEFAULT_LAYOUT,
  selection: null,
  activeSample: "expenseApproval",
  hasDiagram: false,
  lastGeneratedCode: "",
  historyEntries: [],
  currentHistoryId: "",
  currentHistoryTitle: "",
};

let cy = null;
let suppressCodeSync = false;
let viewportResizeTimer = 0;
let codeSyncTimer = 0;
let inlineEditorState = null;
let pendingInlineEditor = null;
let lastCanvasTap = { kind: "", id: "", time: 0 };
let nodeLabelMeasureContext = null;
let connectorDraft = null;

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  configureDocumentReaders();
  populateSamples();
  bindEvents();
  initializeCanvas();
  hydrateVersionHistory();
  hydrateState();
});

function cacheDom() {
  refs.intakeScreen = document.getElementById("intakeScreen");
  refs.workspaceScreen = document.getElementById("workspaceScreen");
  refs.brandHomeLink = document.getElementById("brandHomeLink");
  refs.sampleButtons = document.getElementById("sampleButtons");
  refs.attachmentInput = document.getElementById("attachmentInput");
  refs.uploadAttachmentButton = document.getElementById("uploadAttachmentButton");
  refs.attachmentStatus = document.getElementById("attachmentStatus");
  refs.clearPromptButton = document.getElementById("clearPromptButton");
  refs.promptInput = document.getElementById("promptInput");
  refs.workspacePromptInput = document.getElementById("workspacePromptInput");
  refs.createDiagramButton = document.getElementById("createDiagramButton");
  refs.updateFlowchartButton = document.getElementById("updateFlowchartButton");
  refs.screenTabIntake = document.getElementById("screenTabIntake");
  refs.screenTabWorkspace = document.getElementById("screenTabWorkspace");
  refs.screenTabHistory = document.getElementById("screenTabHistory");
  refs.startOverButton = document.getElementById("startOverButton");
  refs.layoutSelect = document.getElementById("layoutSelect");
  refs.diagramShell = document.getElementById("diagramShell");
  refs.diagramViewport = document.getElementById("diagramViewport");
  refs.diagramCanvas = document.getElementById("diagramCanvas");
  refs.emptyCanvas = document.getElementById("emptyCanvas");
  refs.nodeControlLayer = document.getElementById("nodeControlLayer");
  refs.diagramStatus = document.getElementById("diagramStatus");
  refs.downloadDiagramButton = document.getElementById("downloadDiagramButton");
  refs.viewFullscreenButton = document.getElementById("viewFullscreenButton");
  refs.returnFromFullscreenButton = document.getElementById("returnFromFullscreenButton");
  refs.zoomOutButton = document.getElementById("zoomOutButton");
  refs.zoomInButton = document.getElementById("zoomInButton");
  refs.fitDiagramButton = document.getElementById("fitDiagramButton");
  refs.zoomLevel = document.getElementById("zoomLevel");
  refs.diagramEditorLayer = document.getElementById("diagramEditorLayer");
  refs.inlineNodeEditor = document.getElementById("inlineNodeEditor");
  refs.inlineEdgeEditor = document.getElementById("inlineEdgeEditor");
  refs.codeEditor = document.getElementById("codeEditor");
  refs.codeMeta = document.getElementById("codeMeta");
  refs.codeStatus = document.getElementById("codeStatus");
  refs.updateCodeButton = document.getElementById("updateCodeButton");
  refs.copyCodeButton = document.getElementById("copyCodeButton");
  refs.resetCodeButton = document.getElementById("resetCodeButton");
  refs.historyScreen = document.getElementById("historyScreen");
  refs.historySummary = document.getElementById("historySummary");
  refs.historyEmptyState = document.getElementById("historyEmptyState");
  refs.historyList = document.getElementById("historyList");
  refs.saveVersionButton = document.getElementById("saveVersionButton");
}

function configureDocumentReaders() {
  if (window.pdfjsLib?.GlobalWorkerOptions) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_SRC;
  }
}

function populateSamples() {
  refs.sampleButtons.innerHTML = "";

  Object.entries(SAMPLE_PROMPTS).forEach(([key, sample]) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "sample-chip ghost-button";
    chip.textContent = sample.label;
    chip.dataset.sample = key;
    refs.sampleButtons.appendChild(chip);
  });
}

function bindEvents() {
  refs.brandHomeLink.addEventListener("click", (event) => {
    event.preventDefault();
    resetToStart();
  });

  refs.sampleButtons.addEventListener("click", (event) => {
    const button = event.target.closest("[data-sample]");
    if (!button) {
      return;
    }

    loadSample(button.dataset.sample);
  });

  refs.uploadAttachmentButton.addEventListener("click", () => {
    refs.attachmentInput.click();
  });

  refs.attachmentInput.addEventListener("change", handleAttachmentSelection);

  refs.clearPromptButton.addEventListener("click", () => {
    clearAttachmentState();
    updatePrompt("");
    refs.promptInput.focus();
    setDiagramStatus(
      "Prompt cleared. Add workflow text or upload an attachment to continue."
    );
  });

  refs.promptInput.addEventListener("input", (event) => {
    updatePrompt(event.target.value, "intake");
  });

  refs.workspacePromptInput.addEventListener("input", (event) => {
    updatePrompt(event.target.value, "workspace");
  });

  refs.createDiagramButton.addEventListener("click", () => {
    buildDiagramFromPrompt(true, "Diagram created from the workflow prompt.");
  });

  refs.updateFlowchartButton.addEventListener("click", () => {
    buildDiagramFromPrompt(false, "Flowchart updated from the edited prompt.");
  });

  refs.screenTabIntake.addEventListener("click", () => {
    showScreen("intake");
  });

  refs.screenTabWorkspace.addEventListener("click", () => {
    if (state.hasDiagram) {
      showScreen("workspace");
      return;
    }

    refs.promptInput.focus();
    setDiagramStatus("Create a diagram first to open the workspace.");
  });
  refs.screenTabHistory.addEventListener("click", () => {
    showScreen("history");
  });

  refs.startOverButton.addEventListener("click", resetToStart);
  refs.saveVersionButton.addEventListener("click", saveCurrentVersion);
  refs.historyList.addEventListener("click", handleHistoryListClick);

  refs.layoutSelect.addEventListener("change", () => {
    state.layout = refs.layoutSelect.value || DEFAULT_LAYOUT;
    if (state.graph) {
      state.graph.direction = state.layout;
      commitGraph(state.graph, {
        source: "layout",
        updateCode: true,
        preserveSelection: true,
        successMessage: "Layout updated.",
      });
    } else {
      persistState();
    }
  });

  refs.codeEditor.addEventListener("input", () => {
    if (suppressCodeSync) {
      return;
    }

    state.code = refs.codeEditor.value;
    refs.codeMeta.textContent = "Edited locally";
    setCodeStatus("Code edited. Updating the diagram preview as the code becomes valid.");
    scheduleCodeSyncFromEditor();
    persistState();
  });

  refs.updateCodeButton.addEventListener("click", applyCodeChanges);
  refs.copyCodeButton.addEventListener("click", copyCodeToClipboard);
  refs.resetCodeButton.addEventListener("click", () => {
    if (!state.lastGeneratedCode) {
      return;
    }

    setCodeEditorValue(state.lastGeneratedCode);
    applyCodeChanges();
  });

  refs.downloadDiagramButton.addEventListener("click", downloadDiagram);
  refs.viewFullscreenButton.addEventListener("click", enterDiagramFullscreen);
  refs.returnFromFullscreenButton.addEventListener("click", exitDiagramFullscreen);
  refs.nodeControlLayer.addEventListener("click", handleNodeControlLayerClick);
  refs.zoomInButton.addEventListener("click", () => {
    zoomDiagram(1.15);
  });
  refs.zoomOutButton.addEventListener("click", () => {
    zoomDiagram(1 / 1.15);
  });
  refs.fitDiagramButton.addEventListener("click", () => {
    fitDiagramToViewport();
    setDiagramStatus("Diagram fitted to the available space.", "Fitted");
  });
  refs.inlineNodeEditor.addEventListener("keydown", handleInlineNodeEditorKeydown);
  refs.inlineEdgeEditor.addEventListener("keydown", handleInlineEdgeEditorKeydown);
  refs.inlineNodeEditor.addEventListener("input", positionInlineEditor);
  refs.inlineEdgeEditor.addEventListener("input", positionInlineEditor);
  refs.inlineNodeEditor.addEventListener("blur", handleInlineEditorBlur);
  refs.inlineEdgeEditor.addEventListener("blur", handleInlineEditorBlur);
  refs.diagramViewport.addEventListener("dblclick", handleDiagramViewportDoubleClick);
  document.addEventListener("click", handleDocumentClick);
  document.addEventListener("fullscreenchange", handleFullscreenChange);
  document.addEventListener("webkitfullscreenchange", handleFullscreenChange);
  window.addEventListener("resize", handleViewportResize, { passive: true });
}

function handleDocumentClick(event) {
  void event;
}

function setAddNodeMenuOpen(isOpen) {
  void isOpen;
}

function setNodeActionMenuTarget(nodeId) {
  void nodeId;
}

function commitInlineEditorIfNeeded(cancelFollowUp = false) {
  if (!inlineEditorState) {
    return;
  }

  applyInlineEditorChanges();

  if (cancelFollowUp) {
    pendingInlineEditor = null;
  }
}

function scheduleCodeSyncFromEditor() {
  window.clearTimeout(codeSyncTimer);
  codeSyncTimer = window.setTimeout(() => {
    syncCodeEditorToDiagram();
  }, 260);
}

function syncCodeEditorToDiagram(options = {}) {
  const code = refs.codeEditor.value;
  state.code = code;

  if (!code.trim()) {
    state.hasDiagram = false;
    clearRenderedDiagramSurface();
    refs.codeMeta.textContent = "Edited directly";
    setCodeStatus("Code cleared. The diagram preview is waiting for new flowchart code.");
    setDiagramStatus("The diagram preview is waiting for flowchart code.");
    persistState();
    updateWorkspaceTabAvailability();
    return false;
  }

  const parsed = parseFlowchartCode(code);
  if (!parsed.ok) {
    setCodeStatus(parsed.errors.join(" "), "error");
    setDiagramStatus(
      "The preview stays on the last valid diagram until the code is valid again."
    );
    persistState();
    return false;
  }

  state.layout = parsed.graph.direction;
  refs.layoutSelect.value = state.layout;

  commitGraph(parsed.graph, {
    source: options.source || "code",
    updateCode: false,
    preserveSelection: options.preserveSelection !== false,
    successMessage: options.successMessage || "Diagram refreshed from code.",
  });

  return true;
}

function handleNodeControlLayerClick(event) {
  const actionButton = event.target.closest("[data-canvas-action]");
  if (actionButton) {
    const targetKind = actionButton.dataset.targetKind;
    const targetId = actionButton.dataset.targetId;
    if (!targetKind || !targetId) {
      return;
    }

    commitInlineEditorIfNeeded(true);
    event.stopPropagation();

    if (actionButton.dataset.canvasAction === "add-process") {
      addNodeFromCanvas("process", { kind: targetKind, id: targetId });
      return;
    }

    if (actionButton.dataset.canvasAction === "add-process-before") {
      insertNodeBeforeCanvas("process", targetId);
      return;
    }

    if (actionButton.dataset.canvasAction === "add-decision") {
      addNodeFromCanvas("decision", { kind: targetKind, id: targetId });
      return;
    }

    if (actionButton.dataset.canvasAction === "add-decision-before") {
      insertNodeBeforeCanvas("decision", targetId);
      return;
    }

    if (actionButton.dataset.canvasAction === "connect") {
      startConnectorDraft({
        mode: "create",
        sourceNodeId: targetId,
      });
      return;
    }

    if (actionButton.dataset.canvasAction === "connect-yes") {
      addDecisionBranchFromCanvas(targetId, "Yes");
      return;
    }

    if (actionButton.dataset.canvasAction === "connect-no") {
      addDecisionBranchFromCanvas(targetId, "No");
      return;
    }

    if (actionButton.dataset.canvasAction === "reassign-source") {
      startConnectorDraft({
        mode: "reassign-source",
        edgeId: targetId,
      });
      return;
    }

    if (actionButton.dataset.canvasAction === "reassign-target") {
      startConnectorDraft({
        mode: "reassign-target",
        edgeId: targetId,
      });
      return;
    }

    if (actionButton.dataset.canvasAction === "cancel-connector") {
      cancelConnectorDraft("Connector update canceled.");
      return;
    }

    if (actionButton.dataset.canvasAction === "edit") {
      openInlineEditor({ kind: targetKind, id: targetId, selectAll: true });
      return;
    }

    if (actionButton.dataset.canvasAction === "remove-node") {
      removeNodeFromCanvas(targetId);
      return;
    }

    if (actionButton.dataset.canvasAction === "remove-edge") {
      removeEdgeFromCanvas(targetId);
    }
    return;
  }

  const button = event.target.closest("[data-node-control]");
  if (!button) {
    return;
  }

  const nodeId = button.dataset.nodeId;
  if (!nodeId) {
    return;
  }

  commitInlineEditorIfNeeded(true);
  event.stopPropagation();

  if (button.dataset.nodeControl === "remove") {
    removeNodeFromCanvas(nodeId);
    return;
  }

  addNodeFromCanvas("process", {
    kind: "node",
    id: nodeId,
  });
}

function handleInlineEditorBlur() {
  const blurSession = inlineEditorState
    ? { kind: inlineEditorState.kind, id: inlineEditorState.id }
    : null;

  window.setTimeout(() => {
    if (
      !blurSession ||
      !inlineEditorState ||
      inlineEditorState.kind !== blurSession.kind ||
      inlineEditorState.id !== blurSession.id
    ) {
      return;
    }

    if (
      document.activeElement === refs.inlineNodeEditor ||
      document.activeElement === refs.inlineEdgeEditor
    ) {
      return;
    }

    applyInlineEditorChanges();
  }, 0);
}

function handleInlineNodeEditorKeydown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    hideInlineEditors();
    return;
  }

  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    applyInlineEditorChanges();
  }
}

function handleInlineEdgeEditorKeydown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    hideInlineEditors();
    return;
  }

  if (event.key === "Enter") {
    event.preventDefault();
    applyInlineEditorChanges();
  }
}

function addNodeFromCanvas(nodeType, preferredInsertionPoint = null) {
  if (!state.graph) {
    setDiagramStatus("Create a diagram before adding new shapes.", "Diagram needed");
    return;
  }

  commitInlineEditorIfNeeded(true);
  const graph = cloneGraph(state.graph);
  const insertionPoint = resolveInsertionPoint(graph, preferredInsertionPoint);

  if (!insertionPoint) {
    setAddNodeMenuOpen(false);
    setDiagramStatus(
      "Select a step or connector in the diagram before adding a new shape.",
      "Select a target"
    );
    return;
  }

  hideInlineEditors();
  setAddNodeMenuOpen(false);

  const normalizedType = nodeType === "decision" ? "decision" : "process";
  const newNode = {
    id: nextNodeId(graph, normalizedType === "decision" ? "decision" : "step"),
    label: normalizedType === "decision" ? "New condition" : "New step",
    type: normalizedType,
  };

  graph.nodes.push(newNode);
  const followUp = insertNodeIntoGraph(graph, insertionPoint, newNode);

  if (!followUp.ok) {
    graph.nodes = graph.nodes.filter((node) => node.id !== newNode.id);
    setDiagramStatus(followUp.message, "Select a target");
    return;
  }

  state.selection = { kind: "node", id: newNode.id };
  pendingInlineEditor = {
    kind: "node",
    id: newNode.id,
    selectAll: true,
    next: followUp.nextEditor || null,
  };

  commitGraph(graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: true,
    successMessage:
      normalizedType === "decision"
        ? "Decision added to the diagram."
        : "Process step added to the diagram.",
  });
}

function insertNodeBeforeCanvas(nodeType, targetNodeId) {
  if (!state.graph) {
    setDiagramStatus("Create a diagram before adding new shapes.");
    return;
  }

  commitInlineEditorIfNeeded(true);
  const graph = cloneGraph(state.graph);
  const targetNode = findNode(targetNodeId, graph);
  if (!targetNode) {
    setDiagramStatus("Select a valid step before inserting a new one.");
    return;
  }

  if (targetNode.type === "start") {
    setDiagramStatus("Insert the next step after Start instead of before it.");
    return;
  }

  hideInlineEditors();
  setAddNodeMenuOpen(false);
  const normalizedType = nodeType === "decision" ? "decision" : "process";
  const newNode = {
    id: nextNodeId(graph, normalizedType === "decision" ? "decision" : "step"),
    label: normalizedType === "decision" ? "New condition" : "New step",
    type: normalizedType,
  };

  graph.nodes.push(newNode);
  const followUp = insertNodeBeforeGraph(graph, targetNodeId, newNode);
  if (!followUp.ok) {
    graph.nodes = graph.nodes.filter((node) => node.id !== newNode.id);
    setDiagramStatus(followUp.message);
    return;
  }

  state.selection = { kind: "node", id: newNode.id };
  pendingInlineEditor = {
    kind: "node",
    id: newNode.id,
    selectAll: true,
    next: followUp.nextEditor || null,
  };

  commitGraph(graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: true,
    successMessage:
      normalizedType === "decision"
        ? "Decision inserted before the selected step."
        : "Process step inserted before the selected step.",
  });
}

function startConnectorDraft(draft) {
  if (!state.graph || !draft) {
    return;
  }

  commitInlineEditorIfNeeded(true);
  hideInlineEditors();
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");

  if (draft.mode === "create" && !findNode(draft.sourceNodeId)) {
    return;
  }

  if (
    (draft.mode === "reassign-source" || draft.mode === "reassign-target")
    && !findEdge(draft.edgeId)
  ) {
    return;
  }

  connectorDraft = draft;
  renderNodeControls();
  setDiagramStatus(describeConnectorDraft(draft).status);
}

function cancelConnectorDraft(message = "") {
  if (!connectorDraft) {
    return;
  }

  connectorDraft = null;
  renderNodeControls();
  if (message) {
    setDiagramStatus(message);
  }
}

function completeConnectorDraft(targetNodeId) {
  if (!state.graph || !connectorDraft) {
    return;
  }

  const graph = cloneGraph(state.graph);
  const draft = connectorDraft;
  const targetNode = findNode(targetNodeId, graph);
  if (!targetNode) {
    cancelConnectorDraft("Select a valid target node for that connector.");
    return;
  }

  if (draft.mode === "create") {
    const sourceNode = findNode(draft.sourceNodeId, graph);
    if (!sourceNode) {
      cancelConnectorDraft("The source node is no longer available.");
      return;
    }

    if (sourceNode.id === targetNode.id) {
      setDiagramStatus("Connect this shape to a different node.");
      return;
    }

    if (targetNode.type === "start") {
      setDiagramStatus("Connectors cannot point into Start.");
      return;
    }

    const edgeId = nextEdgeId(graph);
    const added = addUniqueGraphEdge(graph, {
      id: edgeId,
      source: sourceNode.id,
      target: targetNode.id,
      label: "",
    });

    if (!added) {
      setDiagramStatus("That connector already exists or cannot be created there.");
      return;
    }

    connectorDraft = null;
    state.selection = { kind: "edge", id: edgeId };
    pendingInlineEditor =
      sourceNode.type === "decision"
        ? { kind: "edge", id: edgeId, selectAll: true }
        : null;

    commitGraph(graph, {
      source: "canvas",
      updateCode: true,
      preserveSelection: true,
      successMessage: "Connector added to the diagram.",
    });
    return;
  }

  const edge = findEdge(draft.edgeId, graph);
  if (!edge) {
    cancelConnectorDraft("That connector is no longer available.");
    return;
  }

  let nextSource = edge.source;
  let nextTarget = edge.target;

  if (draft.mode === "reassign-source") {
    nextSource = targetNode.id;
  } else {
    nextTarget = targetNode.id;
  }

  if (draft.mode === "reassign-target" && targetNode.type === "start") {
    setDiagramStatus("Connectors cannot point into Start.");
    return;
  }

  if (nextSource === nextTarget) {
    setDiagramStatus("A connector must link two different nodes.");
    return;
  }

  const isDuplicate = graph.edges.some(
    (entry) =>
      entry.id !== edge.id &&
      entry.source === nextSource &&
      entry.target === nextTarget &&
      cleanLabel(entry.label) === cleanLabel(edge.label)
  );

  if (isDuplicate) {
    setDiagramStatus("That connector already exists.");
    return;
  }

  if (nextSource === edge.source && nextTarget === edge.target) {
    cancelConnectorDraft("Connector kept as-is.");
    return;
  }

  edge.source = nextSource;
  edge.target = nextTarget;
  connectorDraft = null;
  state.selection = { kind: "edge", id: edge.id };
  pendingInlineEditor =
    findNode(edge.source, graph)?.type === "decision" && !cleanLabel(edge.label)
      ? { kind: "edge", id: edge.id, selectAll: true }
      : null;

  commitGraph(graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: true,
    successMessage: "Connector reassigned on the diagram.",
  });
}

function addDecisionBranchFromCanvas(nodeId, branchLabel) {
  if (!state.graph) {
    return;
  }

  commitInlineEditorIfNeeded(true);
  const graph = cloneGraph(state.graph);
  const decisionNode = findNode(nodeId, graph);
  const normalizedBranchLabel = cleanLabel(branchLabel) || "Yes";
  if (!decisionNode || decisionNode.type !== "decision") {
    setDiagramStatus("Select a decision node before adding a branch.");
    return;
  }

  if (hasDecisionBranchLabel(nodeId, normalizedBranchLabel, graph)) {
    setDiagramStatus(`This decision already has a ${normalizedBranchLabel.toLowerCase()} branch.`);
    return;
  }

  hideInlineEditors();
  const newNode = {
    id: nextNodeId(graph, "step"),
    label: "New step",
    type: "process",
  };
  graph.nodes.push(newNode);

  const edgeId = nextEdgeId(graph);
  addUniqueGraphEdge(graph, {
    id: edgeId,
    source: nodeId,
    target: newNode.id,
    label: normalizedBranchLabel,
  });
  reorderNodeAfter(graph, newNode.id, nodeId);

  state.selection = { kind: "node", id: newNode.id };
  pendingInlineEditor = {
    kind: "node",
    id: newNode.id,
    selectAll: true,
  };

  commitGraph(graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: true,
    successMessage: `${normalizedBranchLabel} branch added to the diagram.`,
  });
}

function describeConnectorDraft(draft, graph = state.graph) {
  if (!draft) {
    return {
      hint: "",
      status: "",
    };
  }

  if (draft.mode === "create") {
    const sourceNode = findNode(draft.sourceNodeId, graph);
    const sourceLabel = sourceNode ? sourceNode.label : "this node";
    return {
      hint: `Connecting from ${sourceLabel}`,
      status: `Select the target node for the connector from "${sourceLabel}".`,
    };
  }

  const edge = findEdge(draft.edgeId, graph);
  const sourceLabel = findNode(edge?.source, graph)?.label || "source";
  const targetLabel = findNode(edge?.target, graph)?.label || "target";
  const modeLabel =
    draft.mode === "reassign-source" ? "Choose a new source node" : "Choose a new target node";

  return {
    hint: `${modeLabel} for ${sourceLabel} -> ${targetLabel}`,
    status: `${modeLabel} for the selected connector.`,
  };
}

function resolveInsertionPoint(graph, preferredInsertionPoint = null) {
  if (preferredInsertionPoint?.id) {
    if (
      (preferredInsertionPoint.kind === "node" && findNode(preferredInsertionPoint.id, graph)) ||
      (preferredInsertionPoint.kind === "edge" && findEdge(preferredInsertionPoint.id, graph))
    ) {
      return preferredInsertionPoint;
    }
  }

  if (state.selection?.id) {
    if (
      (state.selection.kind === "node" && findNode(state.selection.id, graph)) ||
      (state.selection.kind === "edge" && findEdge(state.selection.id, graph))
    ) {
      return state.selection;
    }
  }

  const defaultInsertionPoint = findDefaultInsertionPoint(graph);
  if (defaultInsertionPoint) {
    state.selection = defaultInsertionPoint;
  }

  return defaultInsertionPoint;
}

function findDefaultInsertionPoint(graph) {
  const endNode = graph.nodes.find((node) => node.type === "end");
  if (endNode) {
    const incomingToEnd = graph.edges.filter((edge) => edge.target === endNode.id);
    if (incomingToEnd.length === 1) {
      return { kind: "edge", id: incomingToEnd[0].id };
    }
  }

  const leafNodes = graph.nodes.filter((node) => {
    if (node.type === "end") {
      return false;
    }

    return !graph.edges.some((edge) => edge.source === node.id);
  });

  if (leafNodes.length === 1) {
    return { kind: "node", id: leafNodes[0].id };
  }

  const lastNonTerminalNode = [...graph.nodes].reverse().find((node) => node.type !== "end");
  if (lastNonTerminalNode) {
    return { kind: "node", id: lastNonTerminalNode.id };
  }

  return null;
}

function insertNodeIntoGraph(graph, insertionPoint, newNode) {
  if (insertionPoint.kind === "edge") {
    const originalEdge = findEdge(insertionPoint.id, graph);
    if (!originalEdge) {
      return {
        ok: false,
        message: "Select a valid connector before adding a new shape.",
      };
    }

    const sourceNode = findNode(originalEdge.source, graph);
    const split = splitEdgeWithNode(graph, originalEdge.id, newNode.id, {
      preserveLabelAtSource: sourceNode?.type === "decision",
    });

    if (!split) {
      return {
        ok: false,
        message: "That connector could not be updated. Try selecting a different point.",
      };
    }

    return {
      ok: true,
      nextEditor:
        newNode.type === "decision"
          ? { kind: "edge", id: split.outgoingEdgeId, selectAll: true }
          : null,
    };
  }

  const selectedNode = findNode(insertionPoint.id, graph);
  if (!selectedNode) {
    return {
      ok: false,
      message: "Select a valid step before adding a new shape.",
    };
  }

  const outgoing = graph.edges.filter((edge) => edge.source === selectedNode.id);
  const incoming = graph.edges.filter((edge) => edge.target === selectedNode.id);

  if (selectedNode.type === "decision") {
    const branchEdgeId = nextEdgeId(graph);
    graph.edges.push({
      id: branchEdgeId,
      source: selectedNode.id,
      target: newNode.id,
      label: "",
    });

    return {
      ok: true,
      nextEditor: { kind: "edge", id: branchEdgeId, selectAll: true },
    };
  }

  if (outgoing.length === 1) {
    const split = splitEdgeWithNode(graph, outgoing[0].id, newNode.id, {
      preserveLabelAtSource: false,
    });

    if (!split) {
      return {
        ok: false,
        message: "The next step could not be inserted there. Try selecting a connector instead.",
      };
    }

    return {
      ok: true,
      nextEditor:
        newNode.type === "decision"
          ? { kind: "edge", id: split.outgoingEdgeId, selectAll: true }
          : null,
    };
  }

  if (outgoing.length === 0) {
    if (selectedNode.type === "end") {
      if (incoming.length !== 1) {
        return {
          ok: false,
          message: "Select the connector leading into End to insert a new shape there.",
        };
      }

      const split = splitEdgeWithNode(graph, incoming[0].id, newNode.id, {
        preserveLabelAtSource: findNode(incoming[0].source, graph)?.type === "decision",
      });

      if (!split) {
        return {
          ok: false,
          message: "The step before End could not be updated. Try selecting the connector instead.",
        };
      }

      return {
        ok: true,
        nextEditor:
          newNode.type === "decision"
            ? { kind: "edge", id: split.outgoingEdgeId, selectAll: true }
            : null,
      };
    }

    const newEdgeId = nextEdgeId(graph);
    graph.edges.push({
      id: newEdgeId,
      source: selectedNode.id,
      target: newNode.id,
      label: "",
    });

    return {
      ok: true,
      nextEditor:
        newNode.type === "decision"
          ? { kind: "edge", id: newEdgeId, selectAll: true }
          : null,
    };
  }

  return {
    ok: false,
    message: "This step branches in multiple directions. Select a specific connector to insert a new shape there.",
  };
}

function insertNodeBeforeGraph(graph, targetNodeId, newNode) {
  const targetNode = findNode(targetNodeId, graph);
  if (!targetNode) {
    return {
      ok: false,
      message: "Select a valid step before inserting a new one.",
    };
  }

  if (targetNode.type === "start") {
    return {
      ok: false,
      message: "Start must stay first. Insert the next step after it instead.",
    };
  }

  const incoming = graph.edges.filter((edge) => edge.target === targetNodeId);
  if (incoming.length !== 1) {
    return {
      ok: false,
      message:
        incoming.length > 1
          ? "This step has multiple incoming connectors. Select a specific connector to insert before it."
          : "This step has no incoming connector to split.",
    };
  }

  const split = splitEdgeWithNode(graph, incoming[0].id, newNode.id, {
    preserveLabelAtSource: findNode(incoming[0].source, graph)?.type === "decision",
  });
  if (!split) {
    return {
      ok: false,
      message: "The selected step could not be updated there.",
    };
  }

  reorderNodeBefore(graph, newNode.id, targetNodeId);

  return {
    ok: true,
    nextEditor:
      newNode.type === "decision"
        ? { kind: "edge", id: split.outgoingEdgeId, selectAll: true }
        : null,
  };
}

function splitEdgeWithNode(graph, edgeId, newNodeId, options = {}) {
  const edgeIndex = graph.edges.findIndex((edge) => edge.id === edgeId);
  if (edgeIndex === -1) {
    return null;
  }

  const edge = graph.edges[edgeIndex];
  graph.edges.splice(edgeIndex, 1);

  const firstEdgeId = nextEdgeId(graph);
  graph.edges.push({
    id: firstEdgeId,
    source: edge.source,
    target: newNodeId,
    label: options.preserveLabelAtSource ? edge.label : "",
  });

  const secondEdgeId = nextEdgeId(graph);
  graph.edges.push({
    id: secondEdgeId,
    source: newNodeId,
    target: edge.target,
    label: options.preserveLabelAtSource ? "" : edge.label,
  });

  return {
    incomingEdgeId: firstEdgeId,
    outgoingEdgeId: secondEdgeId,
  };
}

function removeNodeFromCanvas(nodeId) {
  if (!state.graph) {
    return;
  }

  commitInlineEditorIfNeeded(true);

  const graph = cloneGraph(state.graph);
  const node = findNode(nodeId, graph);
  if (!node || !isCanvasEditableNode(node)) {
    setDiagramStatus("Only process and decision shapes can be removed directly.", "Unavailable");
    return;
  }

  hideInlineEditors();

  const result = removeNodeFromGraph(graph, nodeId);
  if (!result.ok) {
    setDiagramStatus(result.message, "Update blocked");
    return;
  }

  state.selection = result.selection;
  pendingInlineEditor = null;

  commitGraph(graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: Boolean(result.selection),
    successMessage:
      node.type === "decision"
        ? "Decision removed from the diagram."
        : "Process step removed from the diagram.",
  });
}

function removeEdgeFromCanvas(edgeId) {
  if (!state.graph) {
    return;
  }

  commitInlineEditorIfNeeded(true);
  const graph = cloneGraph(state.graph);
  const edge = findEdge(edgeId, graph);
  if (!edge) {
    setDiagramStatus("That connector is no longer available.", "Unavailable");
    return;
  }

  hideInlineEditors();
  graph.edges = graph.edges.filter((entry) => entry.id !== edgeId);
  state.selection = findNode(edge.source, graph)
    ? { kind: "node", id: edge.source }
    : findNode(edge.target, graph)
      ? { kind: "node", id: edge.target }
      : null;
  pendingInlineEditor = null;

  commitGraph(graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: Boolean(state.selection),
    successMessage: "Connector removed from the diagram.",
  });
}

function removeNodeFromGraph(graph, nodeId) {
  const node = findNode(nodeId, graph);
  if (!node || !isCanvasEditableNode(node)) {
    return {
      ok: false,
      message: "That shape cannot be removed directly.",
    };
  }

  const incoming = graph.edges.filter((edge) => edge.target === nodeId);
  const outgoing = graph.edges.filter((edge) => edge.source === nodeId);

  graph.nodes = graph.nodes.filter((entry) => entry.id !== nodeId);
  graph.edges = graph.edges.filter((edge) => edge.source !== nodeId && edge.target !== nodeId);

  if (incoming.length && outgoing.length) {
    incoming.forEach((incomingEdge) => {
      outgoing.forEach((outgoingEdge) => {
        if (incomingEdge.source === outgoingEdge.target) {
          return;
        }

        addUniqueGraphEdge(graph, {
          id: nextEdgeId(graph),
          source: incomingEdge.source,
          target: outgoingEdge.target,
          label: buildBypassEdgeLabel(node, incomingEdge, outgoingEdge, graph),
        });
      });
    });
  }

  const nextSelection =
    incoming.length && findNode(incoming[0].source, graph)
      ? { kind: "node", id: incoming[0].source }
      : outgoing.length && findNode(outgoing[0].target, graph)
        ? { kind: "node", id: outgoing[0].target }
        : null;

  return {
    ok: true,
    selection: nextSelection,
  };
}

function buildBypassEdgeLabel(removedNode, incomingEdge, outgoingEdge, graph) {
  const incomingSourceType = findNode(incomingEdge.source, graph)?.type;

  if (removedNode.type === "decision") {
    return outgoingEdge.label || incomingEdge.label || "";
  }

  if (incomingSourceType === "decision") {
    return incomingEdge.label || outgoingEdge.label || "";
  }

  return outgoingEdge.label || incomingEdge.label || "";
}

function addUniqueGraphEdge(graph, edge) {
  if (
    !edge?.source ||
    !edge?.target ||
    edge.source === edge.target ||
    graph.edges.some(
      (entry) =>
        entry.source === edge.source &&
        entry.target === edge.target &&
        cleanLabel(entry.label) === cleanLabel(edge.label)
    )
  ) {
    return false;
  }

  graph.edges.push({
    id: edge.id || nextEdgeId(graph),
    source: edge.source,
    target: edge.target,
    label: cleanLabel(edge.label),
  });
  return true;
}

function reorderNodeBefore(graph, nodeId, targetNodeId) {
  const nodeIndex = graph.nodes.findIndex((node) => node.id === nodeId);
  const targetIndex = graph.nodes.findIndex((node) => node.id === targetNodeId);
  if (nodeIndex === -1 || targetIndex === -1 || nodeIndex === targetIndex) {
    return;
  }

  const [node] = graph.nodes.splice(nodeIndex, 1);
  const nextTargetIndex = graph.nodes.findIndex((entry) => entry.id === targetNodeId);
  graph.nodes.splice(nextTargetIndex, 0, node);
}

function reorderNodeAfter(graph, nodeId, anchorNodeId) {
  const nodeIndex = graph.nodes.findIndex((node) => node.id === nodeId);
  const anchorIndex = graph.nodes.findIndex((node) => node.id === anchorNodeId);
  if (nodeIndex === -1 || anchorIndex === -1 || nodeIndex === anchorIndex) {
    return;
  }

  const [node] = graph.nodes.splice(nodeIndex, 1);
  const nextAnchorIndex = graph.nodes.findIndex((entry) => entry.id === anchorNodeId);
  graph.nodes.splice(nextAnchorIndex + 1, 0, node);
}

function hasDecisionBranchLabel(nodeId, label, graph = state.graph) {
  const normalizedLabel = cleanLabel(label).toLowerCase();
  if (!normalizedLabel) {
    return false;
  }

  return graph?.edges.some(
    (edge) =>
      edge.source === nodeId &&
      cleanLabel(edge.label).toLowerCase() === normalizedLabel
  );
}

function isCanvasEditableNode(node) {
  return node?.type === "process" || node?.type === "decision";
}

function renderNodeControls() {
  refs.nodeControlLayer.innerHTML = "";

  if (!cy || !state.graph || cy.elements().length === 0) {
    return;
  }

  renderSelectionActionBar();
  positionNodeControls();
}

function positionNodeControls() {
  if (!cy) {
    return;
  }

  positionSelectionActionBar();
}

function renderSelectionActionBar() {
  if (!state.selection || !state.graph) {
    return;
  }

  const targetKind = state.selection.kind === "edge" ? "edge" : "node";
  const targetId = state.selection.id;
  if (!targetId) {
    return;
  }

  if (connectorDraft) {
    const bar = document.createElement("div");
    bar.className = "selection-action-bar is-connector-mode";
    bar.dataset.targetKind = targetKind;
    bar.dataset.targetId = targetId;
    const connectorCopy = describeConnectorDraft(connectorDraft);
    if (connectorCopy.hint) {
      bar.appendChild(createSelectionActionHint(connectorCopy.hint));
    }
    bar.appendChild(
      createSelectionActionButton("Cancel", "cancel-connector", targetKind, targetId)
    );
    refs.nodeControlLayer.appendChild(bar);
    return;
  }

  if (targetKind === "node") {
    const node = findNode(targetId);
    const cyNode = cy?.getElementById(targetId);
    if (!node || !cyNode || cyNode.empty()) {
      return;
    }

    const bar = document.createElement("div");
    bar.className = "selection-action-bar";
    bar.dataset.targetKind = targetKind;
    bar.dataset.targetId = targetId;
    const processInsertAction = node.type === "end" ? "add-process-before" : "add-process";
    const processInsertLabel = node.type === "end" ? "Insert Step" : "+ Step";
    const decisionInsertAction = node.type === "end" ? "add-decision-before" : "add-decision";
    const decisionInsertLabel = node.type === "end" ? "Insert Decision" : "+ Decision";
    if (node.type !== "start" && node.type !== "end") {
      bar.appendChild(
        createSelectionActionButton("Step Before", "add-process-before", targetKind, targetId)
      );
    }
    bar.appendChild(
      createSelectionActionButton(processInsertLabel, processInsertAction, targetKind, targetId)
    );
    bar.appendChild(
      createSelectionActionButton(decisionInsertLabel, decisionInsertAction, targetKind, targetId)
    );
    if (node.type === "decision" && !hasDecisionBranchLabel(targetId, "Yes")) {
      bar.appendChild(createSelectionActionButton("+ Yes", "connect-yes", targetKind, targetId));
    }
    if (node.type === "decision" && !hasDecisionBranchLabel(targetId, "No")) {
      bar.appendChild(createSelectionActionButton("+ No", "connect-no", targetKind, targetId));
    }
    bar.appendChild(createSelectionActionButton("Connect", "connect", targetKind, targetId));
    bar.appendChild(createSelectionActionButton("Edit", "edit", targetKind, targetId));
    if (isCanvasEditableNode(node)) {
      bar.appendChild(
        createSelectionActionButton("Delete", "remove-node", targetKind, targetId, "is-danger")
      );
    }
    refs.nodeControlLayer.appendChild(bar);
    return;
  }

  const edge = findEdge(targetId);
  const cyEdge = cy?.getElementById(targetId);
  if (!edge || !cyEdge || cyEdge.empty()) {
    return;
  }

  const bar = document.createElement("div");
  bar.className = "selection-action-bar is-edge";
  bar.dataset.targetKind = targetKind;
  bar.dataset.targetId = targetId;
  bar.appendChild(createSelectionActionButton("+ Step", "add-process", targetKind, targetId));
  bar.appendChild(createSelectionActionButton("+ Decision", "add-decision", targetKind, targetId));
  bar.appendChild(createSelectionActionButton("Label", "edit", targetKind, targetId));
  bar.appendChild(
    createSelectionActionButton("Move Start", "reassign-source", targetKind, targetId)
  );
  bar.appendChild(
    createSelectionActionButton("Move End", "reassign-target", targetKind, targetId)
  );
  bar.appendChild(
    createSelectionActionButton("Delete Link", "remove-edge", targetKind, targetId, "is-danger")
  );
  refs.nodeControlLayer.appendChild(bar);
}

function createSelectionActionButton(label, action, targetKind, targetId, toneClass = "") {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `selection-action-button${toneClass ? ` ${toneClass}` : ""}`;
  button.textContent = label;
  button.dataset.canvasAction = action;
  button.dataset.targetKind = targetKind;
  button.dataset.targetId = targetId;
  return button;
}

function createSelectionActionHint(label) {
  const hint = document.createElement("span");
  hint.className = "selection-action-hint";
  hint.textContent = label;
  return hint;
}

function positionSelectionActionBar() {
  const bar = refs.nodeControlLayer.querySelector(".selection-action-bar");
  if (!bar || !cy || !state.selection?.id) {
    return;
  }

  const viewportWidth = refs.diagramViewport.clientWidth || 0;
  const viewportHeight = refs.diagramViewport.clientHeight || 0;
  const barWidth = bar.offsetWidth || 0;
  const barHeight = bar.offsetHeight || 0;
  const margin = 14;
  let x = margin;
  let y = margin;

  if (state.selection.kind === "edge") {
    const edge = findEdge(state.selection.id);
    if (!edge) {
      bar.style.display = "none";
      return;
    }

    const midpoint = getRenderedEdgeMidpoint(edge);
    x = midpoint.x - barWidth / 2;
    y = midpoint.y - barHeight - 18;
    if (y < margin) {
      y = midpoint.y + 18;
    }
  } else {
    const cyNode = cy.getElementById(state.selection.id);
    if (!cyNode || cyNode.empty()) {
      bar.style.display = "none";
      return;
    }

    const renderedPosition = cyNode.renderedPosition();
    x = renderedPosition.x - barWidth / 2;
    y = renderedPosition.y - cyNode.renderedHeight() / 2 - barHeight - 14;
    if (y < margin) {
      y = renderedPosition.y + cyNode.renderedHeight() / 2 + 14;
    }
  }

  x = clampNumber(x, margin, Math.max(margin, viewportWidth - barWidth - margin));
  y = clampNumber(y, margin, Math.max(margin, viewportHeight - barHeight - margin));
  bar.style.display = "";
  bar.style.left = `${x}px`;
  bar.style.top = `${y}px`;
}

function openPendingInlineEditor() {
  if (!pendingInlineEditor) {
    return;
  }

  const nextEditor = pendingInlineEditor;
  pendingInlineEditor = null;
  openInlineEditor(nextEditor);
}

function openInlineEditor(target) {
  if (!target?.id || !cy) {
    return;
  }

  setAddNodeMenuOpen(false);
  setSelection({ kind: target.kind === "edge" ? "edge" : "node", id: target.id });

  if (target.kind === "edge") {
    showInlineEdgeEditor(target.id, target);
    return;
  }

  showInlineNodeEditor(target.id, target);
}

function showInlineNodeEditor(nodeId, options = {}) {
  const node = findNode(nodeId);
  const cyNode = cy?.getElementById(nodeId);
  if (!node || !cyNode || cyNode.empty()) {
    return;
  }

  inlineEditorState = {
    kind: "node",
    id: nodeId,
    fallbackValue: node.type === "decision" ? "New condition" : "New step",
    next: options.next || null,
  };

  refs.inlineEdgeEditor.hidden = true;
  refs.inlineNodeEditor.value = node.label;
  refs.inlineNodeEditor.placeholder = inlineEditorState.fallbackValue;
  refs.inlineNodeEditor.classList.toggle("is-terminal", node.type === "start" || node.type === "end");
  refs.inlineNodeEditor.classList.toggle("is-decision", node.type === "decision");
  refs.inlineNodeEditor.hidden = false;
  positionInlineEditor();
  refs.inlineNodeEditor.focus();

  if (options.selectAll) {
    refs.inlineNodeEditor.select();
  } else {
    const cursor = refs.inlineNodeEditor.value.length;
    refs.inlineNodeEditor.setSelectionRange(cursor, cursor);
  }
}

function showInlineEdgeEditor(edgeId, options = {}) {
  const edge = findEdge(edgeId);
  const cyEdge = cy?.getElementById(edgeId);
  if (!edge || !cyEdge || cyEdge.empty()) {
    return;
  }

  inlineEditorState = {
    kind: "edge",
    id: edgeId,
    fallbackValue: "",
    next: options.next || null,
  };

  refs.inlineNodeEditor.hidden = true;
  refs.inlineEdgeEditor.value = edge.label || "";
  refs.inlineEdgeEditor.placeholder = "Condition label";
  refs.inlineEdgeEditor.hidden = false;
  positionInlineEditor();
  refs.inlineEdgeEditor.focus();

  if (options.selectAll) {
    refs.inlineEdgeEditor.select();
  } else {
    const cursor = refs.inlineEdgeEditor.value.length;
    refs.inlineEdgeEditor.setSelectionRange(cursor, cursor);
  }
}

function positionInlineEditor() {
  if (!inlineEditorState || !cy) {
    return;
  }

  if (inlineEditorState.kind === "edge") {
    const edge = findEdge(inlineEditorState.id);
    const cyEdge = cy.getElementById(inlineEditorState.id);
    if (!edge || !cyEdge || cyEdge.empty()) {
      hideInlineEditors();
      return;
    }

    const midpoint = getRenderedEdgeMidpoint(edge);
    const width = clampNumber(
      Math.max((refs.inlineEdgeEditor.value || refs.inlineEdgeEditor.placeholder).length * 8 + 44, 132),
      132,
      260
    );

    refs.inlineEdgeEditor.style.width = `${width}px`;
    refs.inlineEdgeEditor.style.left = `${midpoint.x - width / 2}px`;
    refs.inlineEdgeEditor.style.top = `${midpoint.y - 21}px`;
    return;
  }

  const node = findNode(inlineEditorState.id);
  const cyNode = cy.getElementById(inlineEditorState.id);
  if (!node || !cyNode || cyNode.empty()) {
    hideInlineEditors();
    return;
  }

  const renderedPosition = cyNode.renderedPosition();
  const width = Math.max(120, cyNode.renderedWidth() - 8);
  const height = Math.max(52, cyNode.renderedHeight() - 8);

  refs.inlineNodeEditor.style.width = `${width}px`;
  refs.inlineNodeEditor.style.height = `${height}px`;
  refs.inlineNodeEditor.style.left = `${renderedPosition.x - width / 2}px`;
  refs.inlineNodeEditor.style.top = `${renderedPosition.y - height / 2}px`;
}

function getRenderedEdgeMidpoint(edge) {
  const sourceNode = cy.getElementById(edge.source);
  const targetNode = cy.getElementById(edge.target);
  if (sourceNode.empty() || targetNode.empty()) {
    return {
      x: refs.diagramViewport.clientWidth / 2,
      y: refs.diagramViewport.clientHeight / 2,
    };
  }

  const sourcePosition = sourceNode.renderedPosition();
  const targetPosition = targetNode.renderedPosition();

  return {
    x: (sourcePosition.x + targetPosition.x) / 2,
    y: (sourcePosition.y + targetPosition.y) / 2,
  };
}

function applyInlineEditorChanges() {
  if (!inlineEditorState || !state.graph) {
    return;
  }

  const session = inlineEditorState;
  const nextEditor = session.next || null;
  let updated = false;

  if (session.kind === "node") {
    const node = findNode(session.id);
    if (!node) {
      hideInlineEditors();
      return;
    }

    const nextLabel = cleanLabel(refs.inlineNodeEditor.value) || session.fallbackValue;
    updated = node.label !== nextLabel;
    node.label = nextLabel;
  } else {
    const edge = findEdge(session.id);
    if (!edge) {
      hideInlineEditors();
      return;
    }

    const nextLabel = cleanLabel(refs.inlineEdgeEditor.value);
    updated = edge.label !== nextLabel;
    edge.label = nextLabel;
  }

  hideInlineEditors();
  pendingInlineEditor = nextEditor;

  if (!updated) {
    openPendingInlineEditor();
    return;
  }

  commitGraph(state.graph, {
    source: "canvas",
    updateCode: true,
    preserveSelection: true,
    successMessage:
      session.kind === "node"
        ? "Diagram label updated directly on the canvas."
        : "Connector label updated directly on the canvas.",
  });
}

function hideInlineEditors() {
  inlineEditorState = null;
  lastCanvasTap = { kind: "", id: "", time: 0 };
  refs.inlineNodeEditor.hidden = true;
  refs.inlineEdgeEditor.hidden = true;
  refs.inlineNodeEditor.classList.remove("is-terminal", "is-decision");
  refs.inlineNodeEditor.value = "";
  refs.inlineEdgeEditor.value = "";
  refs.inlineNodeEditor.style.left = "";
  refs.inlineNodeEditor.style.top = "";
  refs.inlineNodeEditor.style.width = "";
  refs.inlineNodeEditor.style.height = "";
  refs.inlineEdgeEditor.style.left = "";
  refs.inlineEdgeEditor.style.top = "";
  refs.inlineEdgeEditor.style.width = "";
}

function handleCanvasElementTap(kind, id, event = null) {
  if (
    inlineEditorState &&
    (inlineEditorState.kind !== kind || inlineEditorState.id !== id)
  ) {
    applyInlineEditorChanges();
  }

  if (connectorDraft) {
    if (kind !== "node") {
      setDiagramStatus("Select a node to finish the connector update.");
      return;
    }

    completeConnectorDraft(id);
    return;
  }

  const now = Date.now();
  const nativeDetail = Number(event?.originalEvent?.detail || 0);
  const isSameSelection = state.selection?.kind === kind && state.selection?.id === id;
  const isDoubleTap =
    nativeDetail >= 2 ||
    (lastCanvasTap.kind === kind &&
      lastCanvasTap.id === id &&
      now - lastCanvasTap.time < 350);

  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  setSelection({ kind, id });

  if (isSameSelection || isDoubleTap) {
    lastCanvasTap = { kind: "", id: "", time: 0 };
    openInlineEditor({ kind, id, selectAll: true });
    return;
  }

  lastCanvasTap = { kind, id, time: now };
}

function handleDiagramViewportDoubleClick(event) {
  if (!state.selection || !cy) {
    return;
  }

  const targetElement = event.target instanceof Element ? event.target : null;
  if (
    targetElement?.closest(".node-control-button") ||
    targetElement?.closest(".selection-action-bar") ||
    targetElement?.closest(".diagram-overlay-tools") ||
    targetElement?.closest(".diagram-return-button") ||
    event.target === refs.inlineNodeEditor ||
    event.target === refs.inlineEdgeEditor
  ) {
    return;
  }

  const selectedKind = state.selection.kind === "edge" ? "edge" : "node";
  const selectedElement = cy.getElementById(state.selection.id);
  if (!selectedElement || selectedElement.empty()) {
    return;
  }

  openInlineEditor({
    kind: selectedKind,
    id: state.selection.id,
    selectAll: true,
  });
}

async function handleAttachmentSelection(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  setAttachmentBusy(true, `Reading ${file.name}...`);

  try {
    const extractedText = await extractTextFromAttachment(file);
    if (!extractedText) {
      throw new Error("No readable text was found in that file. Try another document.");
    }

    updateAttachmentUi();
    updatePrompt(extractedText);
    setAttachmentStatus(
      `Text extracted from ${file.name}. Generating the flowchart now.`,
      "success"
    );
    buildDiagramFromPrompt(true, `Diagram created from ${file.name}.`);
  } catch (error) {
    console.error("Unable to process attachment.", error);
    setAttachmentStatus(
      error?.message ||
        "The uploaded file could not be read. Try a PDF, DOCX, TXT, Markdown, JPEG, or PNG file.",
      "error"
    );
  } finally {
    setAttachmentBusy(false);
    refs.attachmentInput.value = "";
  }
}

function clearAttachmentState(clearStatus = true) {
  refs.attachmentInput.value = "";
  updateAttachmentUi();

  if (clearStatus) {
    setAttachmentStatus("");
  }
}

function updateAttachmentUi() {
  refs.uploadAttachmentButton.textContent = "Upload Attachment";
}

function setAttachmentBusy(isBusy, message = "") {
  refs.uploadAttachmentButton.disabled = isBusy;
  refs.createDiagramButton.disabled = isBusy;
  refs.clearPromptButton.disabled = isBusy;
  refs.promptInput.readOnly = isBusy;

  if (isBusy) {
    refs.uploadAttachmentButton.textContent = "Reading Attachment...";
    if (message) {
      setAttachmentStatus(message);
    }
    return;
  }

  refs.createDiagramButton.disabled = false;
  refs.clearPromptButton.disabled = false;
  refs.promptInput.readOnly = false;
  updateAttachmentUi();
}

function setAttachmentStatus(message, tone = "") {
  refs.attachmentStatus.textContent = message;
  refs.attachmentStatus.classList.remove("is-error", "is-success");

  if (tone === "error") {
    refs.attachmentStatus.classList.add("is-error");
  } else if (tone === "success") {
    refs.attachmentStatus.classList.add("is-success");
  }
}

async function extractTextFromAttachment(file) {
  const kind = detectAttachmentKind(file);

  if (kind === "pdf") {
    return extractTextFromPdf(file);
  }

  if (kind === "docx") {
    return extractTextFromDocx(file);
  }

  if (kind === "markdown") {
    return extractTextFromMarkdownFile(file);
  }

  if (kind === "image") {
    return extractTextFromImage(file);
  }

  if (kind === "text") {
    return extractTextFromPlainFile(file);
  }

  throw new Error("Upload a PDF, DOCX, TXT, Markdown, JPEG, or PNG file.");
}

function detectAttachmentKind(file) {
  const name = String(file?.name || "").toLowerCase();
  const type = String(file?.type || "").toLowerCase();

  if (type === "application/pdf" || name.endsWith(".pdf")) {
    return "pdf";
  }

  if (
    type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    name.endsWith(".docx")
  ) {
    return "docx";
  }

  if (type === "text/markdown" || name.endsWith(".md") || name.endsWith(".markdown")) {
    return "markdown";
  }

  if (type === "image/jpeg" || type === "image/png" || name.endsWith(".jpg") || name.endsWith(".jpeg") || name.endsWith(".png")) {
    return "image";
  }

  if (type.startsWith("text/") || name.endsWith(".txt")) {
    return "text";
  }

  return "";
}

async function extractTextFromPdf(file) {
  if (!window.pdfjsLib?.getDocument) {
    throw new Error("PDF reading is not available right now. Reload the page and try again.");
  }

  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await window.pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(extractPdfPageText(content.items));
  }

  return normalizeExtractedText(pages.join("\n\n"));
}

function extractPdfPageText(items) {
  const lines = [];
  let currentLine = [];
  let currentY = null;

  items.forEach((item) => {
    const text = cleanLabel(item?.str);
    if (!text) {
      return;
    }

    const y = typeof item?.transform?.[5] === "number" ? item.transform[5] : currentY;
    const movedToNewLine =
      currentY !== null && y !== null && Math.abs(y - currentY) > 2.4;

    if ((item.hasEOL || movedToNewLine) && currentLine.length) {
      lines.push(currentLine.join(" "));
      currentLine = [];
    }

    currentLine.push(text);
    currentY = y;
  });

  if (currentLine.length) {
    lines.push(currentLine.join(" "));
  }

  return lines.join("\n");
}

async function extractTextFromDocx(file) {
  if (!window.mammoth?.extractRawText) {
    throw new Error("DOCX reading is not available right now. Reload the page and try again.");
  }

  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  return normalizeExtractedText(result.value);
}

async function extractTextFromPlainFile(file) {
  return normalizeExtractedText(await file.text());
}

async function extractTextFromMarkdownFile(file) {
  return normalizeExtractedText(stripMarkdownSyntax(await file.text()));
}

async function extractTextFromImage(file) {
  if (!window.Tesseract?.recognize) {
    throw new Error("Image reading is not available right now. Reload the page and try again.");
  }

  const result = await window.Tesseract.recognize(file, "eng");
  return normalizeExtractedText(result?.data?.text);
}

function normalizeExtractedText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stripMarkdownSyntax(text) {
  return String(text || "")
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/^#{1,6}\s*/gm, "")
    .replace(/^>\s?/gm, "")
    .replace(/!\[([^\]]*)\]\([^)]+\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/[*_`~]+/g, "");
}

function initializeCanvas() {
  if (typeof cytoscape !== "function") {
    refs.diagramCanvas.innerHTML =
      '<div class="empty-state"><p>The interactive diagram library did not load. Check your internet connection and reload the page.</p></div>';
    setDiagramViewportControlsDisabled(true);
    setDiagramStatus("The interactive canvas could not load.", "Unavailable");
    return;
  }

  cy = cytoscape({
    container: refs.diagramCanvas,
    elements: [],
    wheelSensitivity: 0.18,
    minZoom: 0.35,
    maxZoom: 2.5,
    panningEnabled: true,
    userPanningEnabled: true,
    zoomingEnabled: true,
    userZoomingEnabled: true,
    boxSelectionEnabled: false,
    style: [
      {
        selector: "node",
        style: {
          label: "data(label)",
          shape: "data(shape)",
          width: "data(width)",
          height: "data(height)",
          "background-color": "data(fill)",
          "border-color": "data(stroke)",
          "border-width": 2,
          color: "#243042",
          "font-family": "Public Sans",
          "font-size": "data(fontSize)",
          "font-weight": 600,
          "text-wrap": "wrap",
          "text-max-width": "data(textMaxWidth)",
          "text-valign": "center",
          "text-halign": "center",
          "overlay-opacity": 0,
        },
      },
      {
        selector: "edge",
        style: {
          width: 2,
          "line-color": "#12345B",
          "target-arrow-color": "#12345B",
          "target-arrow-shape": "triangle",
          "curve-style": "round-taxi",
          "taxi-direction": "downward",
          "taxi-turn": 30,
          label: "data(label)",
          "font-family": "Public Sans",
          "font-size": 12,
          color: "#516072",
          "text-background-color": "#FFFFFF",
          "text-background-opacity": 1,
          "text-background-padding": 4,
          "text-border-width": 1,
          "text-border-color": "#E2E8F0",
          "text-border-opacity": 1,
          "text-rotation": "autorotate",
          "text-margin-y": -12,
          "overlay-opacity": 0,
        },
      },
      {
        selector: ":selected",
        style: {
          "underlay-color": "#2F9B95",
          "underlay-opacity": 0.16,
          "underlay-padding": 8,
        },
      },
    ],
  });

  cy.on("tap", "node", (event) => {
    handleCanvasElementTap("node", event.target.id(), event);
  });

  cy.on("tap", "edge", (event) => {
    handleCanvasElementTap("edge", event.target.id(), event);
  });

  cy.on("tap", (event) => {
    if (event.target === cy) {
      if (connectorDraft) {
        cancelConnectorDraft("Connector update canceled.");
        return;
      }

      lastCanvasTap = { kind: "", id: "", time: 0 };
      setAddNodeMenuOpen(false);
      setNodeActionMenuTarget("");
      setSelection(null);
    }
  });

  cy.on("zoom pan", () => {
    updateZoomReadout();
    positionInlineEditor();
    positionNodeControls();
  });
  cy.on("position drag", () => {
    positionInlineEditor();
    positionNodeControls();
  });
  cy.on("dragfreeon", () => {
    positionInlineEditor();
    positionNodeControls();
  });
  setDiagramViewportControlsDisabled(true);
  updateZoomReadout();
}

function normalizeWorkspaceSnapshot(saved = {}) {
  return {
    prompt: typeof saved.prompt === "string" ? saved.prompt : "",
    code: typeof saved.code === "string" ? saved.code : "",
    layout: typeof saved.layout === "string" ? saved.layout : DEFAULT_LAYOUT,
    activeSample:
      typeof saved.activeSample === "string" && SAMPLE_PROMPTS[saved.activeSample]
        ? saved.activeSample
        : "expenseApproval",
    hasDiagram: Boolean(saved.hasDiagram),
    lastGeneratedCode:
      typeof saved.lastGeneratedCode === "string" ? saved.lastGeneratedCode : "",
    currentHistoryId:
      typeof saved.currentHistoryId === "string" ? saved.currentHistoryId : "",
    currentHistoryTitle:
      typeof saved.currentHistoryTitle === "string" ? saved.currentHistoryTitle : "",
  };
}

function createWorkspaceSnapshot() {
  return {
    prompt: state.prompt,
    code: state.code,
    layout: state.layout,
    activeSample: state.activeSample,
    hasDiagram: state.hasDiagram,
    lastGeneratedCode: state.lastGeneratedCode,
  };
}

function hasWorkspaceContent(snapshot) {
  return Boolean(
    cleanLabel(snapshot.prompt)
      || snapshot.code.trim()
      || snapshot.hasDiagram
      || snapshot.lastGeneratedCode.trim()
  );
}

function generateHistoryId() {
  return globalThis.crypto?.randomUUID?.()
    || `history_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function isValidIsoDate(value) {
  return typeof value === "string" && !Number.isNaN(Date.parse(value));
}

function deriveHistoryTitle(rawTitle, snapshot) {
  return cleanLabel(rawTitle) || inferSnapshotTitle(snapshot);
}

function inferSnapshotTitle(snapshot) {
  const prompt = cleanLabel(snapshot.prompt);
  const theme = detectPromptTheme(prompt);
  if (theme) {
    return theme;
  }

  const firstLine = prompt
    .split(/\r?\n/)
    .map((line) => trimSentence(line))
    .find(Boolean);
  if (firstLine) {
    return buildTitleFromPromptLine(firstLine);
  }

  if (snapshot.code.trim()) {
    const parsed = parseFlowchartCode(snapshot.code);
    if (parsed.ok) {
      const firstProcessNode = parsed.graph.nodes.find((node) => node.type === "process");
      if (firstProcessNode) {
        return buildTitleFromPromptLine(firstProcessNode.label);
      }
    }
  }

  return "Workflow Diagram";
}

function normalizeHistoryEntry(entry) {
  if (!entry || typeof entry !== "object") {
    return null;
  }

  const snapshot = normalizeWorkspaceSnapshot(entry.snapshot || entry);
  if (!hasWorkspaceContent(snapshot)) {
    return null;
  }

  const createdAt = isValidIsoDate(entry.createdAt) ? entry.createdAt : new Date().toISOString();
  const updatedAt = isValidIsoDate(entry.updatedAt) ? entry.updatedAt : createdAt;

  return {
    id: typeof entry.id === "string" && entry.id ? entry.id : generateHistoryId(),
    title: deriveHistoryTitle(entry.title, snapshot),
    createdAt,
    updatedAt,
    snapshot,
  };
}

function sortHistoryEntries() {
  state.historyEntries.sort((left, right) => Date.parse(right.updatedAt) - Date.parse(left.updatedAt));
}

function getCurrentHistoryEntry() {
  return state.historyEntries.find((entry) => entry.id === state.currentHistoryId) || null;
}

function hydrateVersionHistory() {
  try {
    const raw = localStorage.getItem(HISTORY_STORAGE_KEY);
    if (!raw) {
      state.historyEntries = [];
      renderVersionHistory();
      return;
    }

    const saved = JSON.parse(raw);
    const entries = Array.isArray(saved?.entries)
      ? saved.entries
      : Array.isArray(saved)
        ? saved
        : [];
    state.historyEntries = entries.map(normalizeHistoryEntry).filter(Boolean);
    sortHistoryEntries();
  } catch (error) {
    console.warn("Unable to restore version history.", error);
    state.historyEntries = [];
  }

  renderVersionHistory();
}

function persistVersionHistory() {
  try {
    localStorage.setItem(
      HISTORY_STORAGE_KEY,
      JSON.stringify({
        entries: state.historyEntries,
      })
    );
  } catch (error) {
    console.warn("Unable to persist version history.", error);
  }
}

function formatHistoryTimestamp(value) {
  try {
    return HISTORY_DATE_FORMATTER.format(new Date(value));
  } catch (error) {
    return "Recently";
  }
}

function createHistoryActionButton(label, className, action, entryId) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `compact-button ${className}`;
  button.textContent = label;
  button.dataset.historyAction = action;
  button.dataset.historyId = entryId;
  return button;
}

function updateSaveVersionButtonState() {
  if (!refs.saveVersionButton) {
    return;
  }

  refs.saveVersionButton.disabled = !hasWorkspaceContent(createWorkspaceSnapshot());
  refs.saveVersionButton.textContent = getCurrentHistoryEntry()
    ? "Update Saved Flowchart"
    : "Save Current Flowchart";
}

function renderVersionHistory() {
  if (!refs.historyList || !refs.historyEmptyState || !refs.historySummary) {
    return;
  }

  refs.historyList.innerHTML = "";
  const hasEntries = state.historyEntries.length > 0;
  refs.historyEmptyState.hidden = hasEntries;
  refs.historyList.hidden = !hasEntries;

  if (!hasEntries) {
    refs.historySummary.textContent = "No saved flowcharts yet.";
    updateSaveVersionButtonState();
    return;
  }

  refs.historySummary.textContent = `${state.historyEntries.length} saved ${
    state.historyEntries.length === 1 ? "flowchart" : "flowcharts"
  }. Open one to continue editing or manage it here.`;

  state.historyEntries.forEach((entry) => {
    const article = document.createElement("article");
    article.className = "history-entry";
    if (entry.id === state.currentHistoryId) {
      article.classList.add("is-current");
    }

    const main = document.createElement("div");
    main.className = "history-entry-main";

    const titleRow = document.createElement("div");
    titleRow.className = "history-entry-title-row";

    const title = document.createElement("h4");
    title.className = "history-entry-title";
    title.textContent = entry.title;
    titleRow.appendChild(title);

    if (entry.id === state.currentHistoryId) {
      const badge = document.createElement("span");
      badge.className = "history-entry-badge";
      badge.textContent = "Current";
      titleRow.appendChild(badge);
    }

    const meta = document.createElement("p");
    meta.className = "history-entry-meta";
    meta.textContent = `Last updated ${formatHistoryTimestamp(entry.updatedAt)}`;

    main.appendChild(titleRow);
    main.appendChild(meta);

    const actions = document.createElement("div");
    actions.className = "history-entry-actions";
    actions.appendChild(createHistoryActionButton("Open", "secondary-button", "open", entry.id));
    actions.appendChild(createHistoryActionButton("Rename", "ghost-button", "rename", entry.id));
    actions.appendChild(createHistoryActionButton("Duplicate", "ghost-button", "duplicate", entry.id));
    actions.appendChild(createHistoryActionButton("Delete", "ghost-button danger-button", "delete", entry.id));

    article.appendChild(main);
    article.appendChild(actions);
    refs.historyList.appendChild(article);
  });

  updateSaveVersionButtonState();
}

function hydrateState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      loadSample(state.activeSample);
      return;
    }

    applyWorkspaceSnapshot(JSON.parse(raw), {
      source: "hydrate",
      useSampleFallback: true,
      successMessage: "Restored saved diagram.",
      emptyDiagramMessage: "Restored the saved draft.",
      emptyCodeMessage: "Saved draft restored. Continue editing and save a version anytime.",
    });
  } catch (error) {
    console.warn("Unable to restore saved app state.", error);
    loadSample("expenseApproval");
  }
}

function persistState() {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        prompt: state.prompt,
        code: state.code,
        layout: state.layout,
        activeSample: state.activeSample,
        hasDiagram: state.hasDiagram,
        lastGeneratedCode: state.lastGeneratedCode,
        currentHistoryId: state.currentHistoryId,
        currentHistoryTitle: state.currentHistoryTitle,
      })
    );
  } catch (error) {
    console.warn("Unable to persist app state.", error);
  }

  updateSaveVersionButtonState();
}

function detachCurrentHistoryContext() {
  state.currentHistoryId = "";
  state.currentHistoryTitle = "";
}

function clearRenderedDiagramSurface() {
  window.clearTimeout(codeSyncTimer);
  state.graph = null;
  state.selection = null;
  pendingInlineEditor = null;
  connectorDraft = null;
  lastCanvasTap = { kind: "", id: "", time: 0 };
  hideInlineEditors();
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  refs.emptyCanvas.hidden = false;
  refs.codeMeta.textContent = "Editable and synced";
  if (cy) {
    cy.elements().remove();
    cy.elements().unselect();
    cy.resize();
  }
  setDiagramViewportControlsDisabled(true);
  updateZoomReadout();
}

function applyWorkspaceSnapshot(snapshot, options = {}) {
  const normalized = normalizeWorkspaceSnapshot(snapshot);
  state.prompt = normalized.prompt;
  state.code = normalized.code;
  state.layout = normalized.layout;
  state.activeSample = normalized.activeSample;
  state.hasDiagram = normalized.hasDiagram;
  state.lastGeneratedCode = normalized.lastGeneratedCode;
  state.currentHistoryId = normalized.currentHistoryId;
  state.currentHistoryTitle = normalized.currentHistoryTitle;

  highlightActiveSample(state.activeSample);
  refs.layoutSelect.value = state.layout;
  updatePrompt(
    options.useSampleFallback
      ? state.prompt || SAMPLE_PROMPTS[state.activeSample].text
      : state.prompt
  );
  setCodeEditorValue(state.code);

  if (state.hasDiagram && state.code.trim()) {
    showScreen("workspace", false);
    const parsed = parseFlowchartCode(state.code);
    if (parsed.ok) {
      commitGraph(parsed.graph, {
        source: options.source || "history",
        updateCode: false,
        preserveSelection: false,
        successMessage: options.successMessage || "Loaded saved flowchart.",
      });
    } else {
      clearRenderedDiagramSurface();
      showScreen("workspace", false);
      setCodeStatus(parsed.errors.join(" "), "error");
      setDiagramStatus("Saved code needs attention before it can render.", "Needs review");
      persistState();
    }
  } else {
    clearRenderedDiagramSurface();
    showScreen(options.screen || "intake", false);
    setCodeStatus(
      options.emptyCodeMessage
        || "Draft restored. Continue editing and create the diagram when ready.",
      "success"
    );
    setDiagramStatus(
      options.emptyDiagramMessage || "Draft restored. Continue editing whenever you're ready.",
      "Loaded"
    );
    persistState();
  }

  renderVersionHistory();
}

function saveCurrentVersion() {
  const snapshot = createWorkspaceSnapshot();
  if (!hasWorkspaceContent(snapshot)) {
    setDiagramStatus("Add content or create a diagram before saving to version history.", "Nothing to save");
    return;
  }

  const timestamp = new Date().toISOString();
  const currentEntry = getCurrentHistoryEntry();
  const title = cleanLabel(state.currentHistoryTitle) || inferDownloadTitle();

  if (currentEntry) {
    currentEntry.title = title;
    currentEntry.updatedAt = timestamp;
    currentEntry.snapshot = snapshot;
    state.currentHistoryTitle = currentEntry.title;
  } else {
    const entry = {
      id: generateHistoryId(),
      title,
      createdAt: timestamp,
      updatedAt: timestamp,
      snapshot,
    };
    state.historyEntries.unshift(entry);
    state.currentHistoryId = entry.id;
    state.currentHistoryTitle = entry.title;
  }

  sortHistoryEntries();
  persistVersionHistory();
  persistState();
  renderVersionHistory();
  setCodeStatus("Flowchart saved to version history.", "success");
  setDiagramStatus("Flowchart saved to version history.", "Saved");
}

function handleHistoryListClick(event) {
  const button = event.target.closest("[data-history-action]");
  if (!button) {
    return;
  }

  const { historyAction: action, historyId: entryId } = button.dataset;
  if (!entryId) {
    return;
  }

  if (action === "open") {
    openHistoryEntry(entryId);
    return;
  }

  if (action === "rename") {
    renameHistoryEntry(entryId);
    return;
  }

  if (action === "duplicate") {
    duplicateHistoryEntry(entryId);
    return;
  }

  if (action === "delete") {
    deleteHistoryEntry(entryId);
  }
}

function openHistoryEntry(entryId) {
  const entry = state.historyEntries.find((item) => item.id === entryId);
  if (!entry) {
    return;
  }

  applyWorkspaceSnapshot(
    {
      ...entry.snapshot,
      currentHistoryId: entry.id,
      currentHistoryTitle: entry.title,
    },
    {
      source: "history",
      useSampleFallback: false,
      successMessage: "Saved flowchart loaded.",
      emptyDiagramMessage: "Saved draft loaded.",
      emptyCodeMessage: "Saved draft loaded. Continue editing and save again anytime.",
      screen: entry.snapshot.hasDiagram && entry.snapshot.code.trim() ? "workspace" : "intake",
    }
  );
}

function renameHistoryEntry(entryId) {
  const entry = state.historyEntries.find((item) => item.id === entryId);
  if (!entry) {
    return;
  }

  const nextTitle = cleanLabel(window.prompt("Rename this saved flowchart:", entry.title));
  if (!nextTitle) {
    return;
  }

  entry.title = nextTitle;
  entry.updatedAt = new Date().toISOString();
  if (state.currentHistoryId === entry.id) {
    state.currentHistoryTitle = entry.title;
    persistState();
  }

  sortHistoryEntries();
  persistVersionHistory();
  renderVersionHistory();
}

function duplicateHistoryEntry(entryId) {
  const entry = state.historyEntries.find((item) => item.id === entryId);
  if (!entry) {
    return;
  }

  const timestamp = new Date().toISOString();
  state.historyEntries.unshift({
    id: generateHistoryId(),
    title: `${entry.title} Copy`,
    createdAt: timestamp,
    updatedAt: timestamp,
    snapshot: {
      ...entry.snapshot,
    },
  });

  sortHistoryEntries();
  persistVersionHistory();
  renderVersionHistory();
}

function deleteHistoryEntry(entryId) {
  const entry = state.historyEntries.find((item) => item.id === entryId);
  if (!entry) {
    return;
  }

  const confirmed = window.confirm(`Delete "${entry.title}" from version history?`);
  if (!confirmed) {
    return;
  }

  state.historyEntries = state.historyEntries.filter((item) => item.id !== entryId);
  if (state.currentHistoryId === entryId) {
    detachCurrentHistoryContext();
    persistState();
  }

  persistVersionHistory();
  renderVersionHistory();
}

function updatePrompt(value, source) {
  state.prompt = value || "";

  if (source !== "intake") {
    refs.promptInput.value = state.prompt;
  }

  if (source !== "workspace") {
    refs.workspacePromptInput.value = state.prompt;
  }

  persistState();
}

function loadSample(sampleKey) {
  const sample = SAMPLE_PROMPTS[sampleKey] || SAMPLE_PROMPTS.expenseApproval;
  window.clearTimeout(codeSyncTimer);
  detachCurrentHistoryContext();
  state.activeSample = sampleKey in SAMPLE_PROMPTS ? sampleKey : "expenseApproval";
  pendingInlineEditor = null;
  connectorDraft = null;
  lastCanvasTap = { kind: "", id: "", time: 0 };
  hideInlineEditors();
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  clearAttachmentState();
  highlightActiveSample(state.activeSample);
  updatePrompt(sample.text);
  setDiagramStatus("Sample prompt loaded. Click Create Diagram to build it.");
}

function highlightActiveSample(activeKey) {
  refs.sampleButtons.querySelectorAll("[data-sample]").forEach((button) => {
    const isActive = button.dataset.sample === activeKey;
    button.classList.toggle("secondary-button", isActive);
    button.classList.toggle("ghost-button", !isActive);
  });
}

function showScreen(screenName, scroll = true) {
  const showWorkspace = screenName === "workspace";
  const showHistory = screenName === "history";
  refs.intakeScreen.hidden = showWorkspace || showHistory;
  refs.workspaceScreen.hidden = !showWorkspace;
  refs.historyScreen.hidden = !showHistory;
  updateWorkspaceTabAvailability();
  updateScreenTabs(screenName);

  if ((showWorkspace || showHistory) && scroll) {
    const targetScreen = showWorkspace ? refs.workspaceScreen : refs.historyScreen;
    window.requestAnimationFrame(() => {
      targetScreen.scrollIntoView({ behavior: "smooth", block: "start" });
      if (showWorkspace) {
        refreshDiagramViewport(true);
      }
    });
    return;
  }

  if (showWorkspace) {
    window.requestAnimationFrame(() => {
      refreshDiagramViewport(true);
    });
  }

  if (showHistory) {
    renderVersionHistory();
  }
}

function updateScreenTabs(screenName) {
  refs.screenTabIntake.classList.toggle("is-active", screenName === "intake");
  refs.screenTabWorkspace.classList.toggle("is-active", screenName === "workspace");
  refs.screenTabHistory.classList.toggle("is-active", screenName === "history");
}

function updateWorkspaceTabAvailability() {
  const enabled = Boolean(state.hasDiagram);
  refs.screenTabWorkspace.disabled = !enabled;
}

function buildDiagramFromPrompt(showWorkspaceAfterBuild, successMessage) {
  const prompt = state.prompt.trim();
  if (!prompt) {
    setDiagramStatus("Add workflow text or load a sample prompt before creating the diagram.", "Prompt needed");
    const promptField = refs.workspaceScreen.hidden ? refs.promptInput : refs.workspacePromptInput;
    promptField.focus();
    return;
  }

  window.clearTimeout(codeSyncTimer);
  pendingInlineEditor = null;
  connectorDraft = null;
  lastCanvasTap = { kind: "", id: "", time: 0 };
  hideInlineEditors();
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  const graph = parsePromptToGraph(prompt);
  graph.direction = refs.layoutSelect.value || state.layout || DEFAULT_LAYOUT;

  commitGraph(graph, {
    source: "prompt",
    updateCode: true,
    preserveSelection: false,
    successMessage: successMessage || "Diagram created from the workflow prompt.",
  });

  if (showWorkspaceAfterBuild) {
    showScreen("workspace");
  }
}

function resetToStart() {
  window.clearTimeout(viewportResizeTimer);
  window.clearTimeout(codeSyncTimer);
  if (isDiagramFullscreen()) {
    void exitDiagramFullscreen();
  }
  state.graph = null;
  state.code = "";
  state.prompt = "";
  state.selection = null;
  state.layout = DEFAULT_LAYOUT;
  state.hasDiagram = false;
  state.lastGeneratedCode = "";
  detachCurrentHistoryContext();
  pendingInlineEditor = null;
  connectorDraft = null;
  lastCanvasTap = { kind: "", id: "", time: 0 };

  refs.layoutSelect.value = DEFAULT_LAYOUT;
  clearAttachmentState();
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  hideInlineEditors();
  updatePrompt("");
  setCodeEditorValue("");
  refs.codeMeta.textContent = "Editable and synced";
  setDiagramViewportControlsDisabled(true);
  refs.emptyCanvas.hidden = false;

  if (cy) {
    cy.elements().remove();
    cy.resize();
  }

  setSelection(null);
  updateZoomReadout();
  setCodeStatus("Start with a new prompt or load a sample to create a diagram.");
  setDiagramStatus("Start with a new prompt or load a sample to create a diagram.", "Ready");
  persistState();
  renderVersionHistory();
  showScreen("intake");
  refs.promptInput.focus();
}

function parsePromptToGraph(text) {
  const preparedText = preprocessPromptText(text);
  return looksLikeArrowFlow(preparedText)
    ? parseArrowFlow(preparedText)
    : parseNarrativeFlow(preparedText);
}

function looksLikeArrowFlow(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return false;
  }

  const arrowLines = lines.filter((line) => /--?>|<--?/.test(line)).length;
  return arrowLines >= Math.ceil(lines.length * 0.7);
}

function commitGraph(graph, options = {}) {
  const source = options.source || "graph";
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  connectorDraft = null;
  state.graph = normalizeGraph(graph);
  state.layout = state.graph.direction || state.layout || DEFAULT_LAYOUT;
  refs.layoutSelect.value = state.layout;

  if (options.updateCode) {
    state.code = graphToCode(state.graph);
    state.lastGeneratedCode = state.code;
    setCodeEditorValue(state.code);
    refs.codeMeta.textContent = source === "prompt" ? "Generated from prompt" : "Synced from diagram";
  } else {
    state.code = refs.codeEditor.value;
    refs.codeMeta.textContent = source === "code" ? "Edited directly" : "Editable and synced";
  }

  renderGraph(state.graph, options.preserveSelection);
  state.hasDiagram = true;
  persistState();
  setCodeStatus(options.successMessage || "Diagram synced successfully.", "success");
  setDiagramStatus(options.successMessage || "Diagram synced successfully.", "Synced");
}

function renderGraph(graph, preserveSelection) {
  if (!cy) {
    return;
  }

  const elements = graphToElements(graph);
  setDiagramViewportControlsDisabled(elements.length === 0);
  refs.emptyCanvas.hidden = Boolean(elements.length);
  cy.elements().remove();
  cy.add(elements);
  cy.resize();
  renderNodeControls();

  if (elements.length) {
    runLayout();
  } else {
    updateZoomReadout();
    positionNodeControls();
    openPendingInlineEditor();
  }

  if (preserveSelection && state.selection) {
    const exists = state.selection.kind === "node" ? findNode(state.selection.id) : findEdge(state.selection.id);
    setSelection(exists ? state.selection : null);
  } else {
    setSelection(null);
  }
}

function runLayout() {
  if (!cy || cy.elements().length === 0) {
    positionNodeControls();
    openPendingInlineEditor();
    return;
  }

  const direction = state.layout === "LR" ? "LR" : "TB";
  const layout = cy.layout({
    name: "dagre",
    rankDir: direction,
    nodeSep: 55,
    edgeSep: 30,
    rankSep: direction === "LR" ? 120 : 95,
    fit: true,
    padding: 36,
    animate: true,
    animationDuration: 220,
  });

  cy.one("layoutstop", () => {
    fitDiagramToViewport();
    window.requestAnimationFrame(() => {
      positionInlineEditor();
      positionNodeControls();
      openPendingInlineEditor();
    });
  });

  layout.run();
}

function refreshDiagramViewport(fitToView = false) {
  if (!cy) {
    updateZoomReadout();
    return;
  }

  cy.resize();
  if (fitToView && cy.elements().length) {
    fitDiagramToViewport();
    return;
  }

  updateZoomReadout();
  positionInlineEditor();
  positionNodeControls();
}

function handleViewportResize() {
  if (!cy || refs.workspaceScreen.hidden) {
    return;
  }

  window.clearTimeout(viewportResizeTimer);
  viewportResizeTimer = window.setTimeout(() => {
    refreshDiagramViewport(true);
  }, 140);
}

function isDiagramFullscreen() {
  return (
    document.fullscreenElement === refs.diagramShell ||
    document.webkitFullscreenElement === refs.diagramShell
  );
}

function updateFullscreenUi() {
  const active = isDiagramFullscreen();
  refs.diagramShell.classList.toggle("is-fullscreen", active);
  refs.returnFromFullscreenButton.hidden = !active;
  refs.viewFullscreenButton.hidden = active;
}

async function enterDiagramFullscreen() {
  if (!state.hasDiagram || !refs.diagramShell) {
    return;
  }

  commitInlineEditorIfNeeded();

  const requestFullscreen =
    refs.diagramShell.requestFullscreen?.bind(refs.diagramShell) ||
    refs.diagramShell.webkitRequestFullscreen?.bind(refs.diagramShell);

  if (!requestFullscreen) {
    setDiagramStatus("Fullscreen mode is not available in this browser.", "Unavailable");
    return;
  }

  try {
    await requestFullscreen();
  } catch (error) {
    console.warn("Unable to enter fullscreen mode.", error);
    setDiagramStatus("Fullscreen mode could not be opened in this browser.", "Unavailable");
  }
}

async function exitDiagramFullscreen() {
  commitInlineEditorIfNeeded();

  if (!isDiagramFullscreen()) {
    updateFullscreenUi();
    return;
  }

  const exitFullscreen =
    document.exitFullscreen?.bind(document) ||
    document.webkitExitFullscreen?.bind(document);

  if (!exitFullscreen) {
    updateFullscreenUi();
    return;
  }

  try {
    await exitFullscreen();
  } catch (error) {
    console.warn("Unable to exit fullscreen mode.", error);
    setDiagramStatus("Fullscreen mode could not be closed automatically.", "Unavailable");
  }
}

function handleFullscreenChange() {
  updateFullscreenUi();

  window.requestAnimationFrame(() => {
    refreshDiagramViewport(true);
  });

  if (!state.hasDiagram) {
    return;
  }

  if (isDiagramFullscreen()) {
    setDiagramStatus("Fullscreen view enabled. Double-click any label to edit it inline.", "Fullscreen");
    return;
  }

  setDiagramStatus("Returned to the standard diagram view.", "Preview");
}

function zoomDiagram(multiplier) {
  if (!cy || cy.elements().length === 0) {
    return;
  }

  const nextZoom = clampZoom(cy.zoom() * multiplier);
  cy.zoom({
    level: nextZoom,
    renderedPosition: {
      x: refs.diagramViewport.clientWidth / 2,
      y: refs.diagramViewport.clientHeight / 2,
    },
  });
  updateZoomReadout();
}

function fitDiagramToViewport() {
  if (!cy || cy.elements().length === 0) {
    updateZoomReadout();
    return;
  }

  cy.fit(cy.elements(), 48);
  updateZoomReadout();
  positionInlineEditor();
  positionNodeControls();
}

function clampZoom(level) {
  if (!cy) {
    return level;
  }

  return Math.max(cy.minZoom(), Math.min(cy.maxZoom(), level));
}

function setDiagramViewportControlsDisabled(disabled) {
  refs.downloadDiagramButton.disabled = disabled;
  refs.viewFullscreenButton.disabled = disabled;
  refs.zoomOutButton.disabled = disabled;
  refs.zoomInButton.disabled = disabled;
  refs.fitDiagramButton.disabled = disabled;

  if (disabled) {
    setAddNodeMenuOpen(false);
    setNodeActionMenuTarget("");
    refs.nodeControlLayer.innerHTML = "";
  }
}

function updateZoomReadout() {
  if (!refs.zoomLevel) {
    return;
  }

  if (!cy || cy.elements().length === 0) {
    refs.zoomLevel.textContent = "100%";
    return;
  }

  refs.zoomLevel.textContent = `${Math.round(cy.zoom() * 100)}%`;
}

function graphToElements(graph) {
  const nodes = graph.nodes.map((node) => {
    const theme = NODE_COLORS[node.type] || NODE_COLORS.process;
    const dimensions = measureNode(node.label, node.type);

    return {
      group: "nodes",
      data: {
        id: node.id,
        label: dimensions.label,
        shape: theme.shape,
        fill: theme.fill,
        stroke: theme.stroke,
        width: dimensions.width,
        height: dimensions.height,
        fontSize: dimensions.fontSize,
        textMaxWidth: dimensions.textMaxWidth,
      },
    };
  });

  const edges = graph.edges.map((edge) => ({
    group: "edges",
    data: {
      id: edge.id,
      source: edge.source,
      target: edge.target,
      label: edge.label || "",
    },
  }));

  return [...nodes, ...edges];
}

function measureNode(label, type) {
  const text = cleanLabel(label) || "Step";
  const config = getNodeTypography(text, type);
  const context = getNodeLabelMeasureContext(config.fontSize);
  const lines = wrapLabelForCanvas(text, config.maxTextWidth, context);
  const longestLineWidth = Math.max(
    ...lines.map((line) => Math.ceil(context.measureText(line).width)),
    config.minTextWidth
  );
  const lineBlockHeight = Math.max(
    config.minContentHeight,
    Math.ceil(lines.length * config.lineHeight)
  );

  if (type === "decision") {
    const height = clampNumber(lineBlockHeight + config.paddingY * 2, 110, 240);
    const width = clampNumber(
      Math.max(longestLineWidth + config.paddingX * 2 + 30, height + 36),
      184,
      320
    );

    return {
      label: lines.join("\n"),
      width,
      height,
      fontSize: config.fontSize,
      textMaxWidth: Math.max(104, width - 92),
    };
  }

  if (type === "start" || type === "end") {
    const height = clampNumber(lineBlockHeight + config.paddingY * 2, 82, 138);
    const width = clampNumber(
      Math.max(longestLineWidth + config.paddingX * 2, height + 12),
      98,
      188
    );

    return {
      label: lines.join("\n"),
      width,
      height,
      fontSize: config.fontSize,
      textMaxWidth: Math.max(58, width - 34),
    };
  }

  const width = clampNumber(longestLineWidth + config.paddingX * 2, 184, 320);
  const height = clampNumber(lineBlockHeight + config.paddingY * 2, 82, 240);

  return {
    label: lines.join("\n"),
    width,
    height,
    fontSize: config.fontSize,
    textMaxWidth: Math.max(120, width - 42),
  };
}

function setSelection(selection) {
  state.selection = selection;
  if (!cy) {
    return;
  }

  cy.elements().unselect();

  if (!selection) {
    lastCanvasTap = { kind: "", id: "", time: 0 };
    renderNodeControls();
    persistState();
    return;
  }

  const cyElement = cy.getElementById(selection.id);
  if (cyElement) {
    cyElement.select();
  } else {
    state.selection = null;
  }

  renderNodeControls();
  persistState();
}

function applyCodeChanges() {
  window.clearTimeout(codeSyncTimer);
  pendingInlineEditor = null;
  connectorDraft = null;
  lastCanvasTap = { kind: "", id: "", time: 0 };
  hideInlineEditors();
  setAddNodeMenuOpen(false);
  setNodeActionMenuTarget("");
  syncCodeEditorToDiagram({
    source: "code",
    preserveSelection: true,
    successMessage: "Diagram refreshed from code.",
  });
}

function parseFlowchartCode(code) {
  const lines = String(code).split(/\r?\n/);
  const graph = {
    direction: DEFAULT_LAYOUT,
    nodes: [],
    edges: [],
  };
  const nodesById = new Map();
  const typeHints = new Map();
  const errors = [];
  let sawHeader = false;

  lines.forEach((rawLine, index) => {
    const line = rawLine.trim();
    if (!line || line.startsWith("%%")) {
      return;
    }

    const stripped = line.replace(/;$/, "");
    const headerMatch = stripped.match(/^(?:flowchart|graph)\s+(TB|TD|LR)$/i);
    if (headerMatch) {
      graph.direction = headerMatch[1].toUpperCase() === "TD" ? "TB" : headerMatch[1].toUpperCase();
      sawHeader = true;
      return;
    }

    if (/^classDef\b/i.test(stripped)) {
      return;
    }

    const classMatch = stripped.match(/^class\s+(.+?)\s+([A-Za-z][\w-]*)$/i);
    if (classMatch) {
      classMatch[1]
        .split(",")
        .map((token) => token.trim())
        .filter(Boolean)
        .forEach((id) => typeHints.set(id, classNameToNodeType(classMatch[2])));
      return;
    }

    if (parseEdgeLine(stripped, graph, nodesById)) {
      return;
    }

    const nodeToken = parseNodeToken(stripped);
    if (nodeToken) {
      upsertNode(graph.nodes, nodesById, nodeToken);
      return;
    }

    errors.push(`Line ${index + 1} could not be interpreted.`);
  });

  typeHints.forEach((type, id) => {
    const node = nodesById.get(id);
    if (node && type) {
      node.type = type;
    }
  });

  if (!sawHeader) {
    errors.unshift("Start the code with `flowchart TB` or `flowchart LR`.");
  }

  if (errors.length) {
    return { ok: false, errors };
  }

  graph.nodes.forEach((node) => {
    if (!node.label) {
      node.label = humanizeId(node.id);
    }
  });

  return { ok: true, graph: normalizeGraph(graph) };
}

function parseEdgeLine(line, graph, nodesById) {
  const labeled = line.match(/^(.*?)\s*-->\|(.+?)\|\s*(.*?)$/);
  const plain = line.match(/^(.*?)\s*-->\s*(.*?)$/);
  const match = labeled || plain;

  if (!match) {
    return false;
  }

  const sourceToken = parseEndpoint(match[1].trim());
  const targetToken = parseEndpoint(match[labeled ? 3 : 2].trim());
  const label = labeled ? cleanLabel(match[2]) : "";

  if (!sourceToken || !targetToken) {
    return true;
  }

  upsertNode(graph.nodes, nodesById, sourceToken);
  upsertNode(graph.nodes, nodesById, targetToken);

  if (
    !graph.edges.some(
      (edge) =>
        edge.source === sourceToken.id &&
        edge.target === targetToken.id &&
        edge.label === label
    )
  ) {
    graph.edges.push({
      id: nextEdgeId(graph),
      source: sourceToken.id,
      target: targetToken.id,
      label,
    });
  }

  return true;
}

function parseEndpoint(token) {
  return parseNodeToken(token) || parseBareIdToken(token);
}

function parseBareIdToken(token) {
  const trimmed = token.trim();
  if (!/^[A-Za-z][\w-]*$/.test(trimmed)) {
    return null;
  }

  return {
    id: trimmed,
    label: humanizeId(trimmed),
    type: "process",
  };
}

function parseNodeToken(token) {
  const trimmed = token.trim().replace(/;$/, "");
  const terminalMatch = trimmed.match(/^([A-Za-z][\w-]*)\(\[(.+)\]\)$/);
  if (terminalMatch) {
    return {
      id: terminalMatch[1],
      label: decodeCodeLabel(terminalMatch[2]),
      type: inferTerminalType(decodeCodeLabel(terminalMatch[2])),
    };
  }

  const decisionMatch = trimmed.match(/^([A-Za-z][\w-]*)\{(.+)\}$/);
  if (decisionMatch) {
    return {
      id: decisionMatch[1],
      label: decodeCodeLabel(decisionMatch[2]),
      type: "decision",
    };
  }

  const processMatch = trimmed.match(/^([A-Za-z][\w-]*)\[(.+)\]$/);
  if (processMatch) {
    return {
      id: processMatch[1],
      label: decodeCodeLabel(processMatch[2]),
      type: inferNodeType(processMatch[2]),
    };
  }

  return null;
}

function decodeCodeLabel(label) {
  return cleanLabel(String(label).replace(/^"(.*)"$/, "$1"));
}

function classNameToNodeType(className) {
  const key = String(className).trim().toLowerCase();
  if (key === "terminal" || key === "start" || key === "end") {
    return "start";
  }

  if (key === "decision") {
    return "decision";
  }

  if (key === "process") {
    return "process";
  }

  return "";
}

function graphToCode(graph) {
  const terminalIds = graph.nodes
    .filter((node) => node.type === "start" || node.type === "end")
    .map((node) => node.id);
  const decisionIds = graph.nodes.filter((node) => node.type === "decision").map((node) => node.id);
  const processIds = graph.nodes
    .filter((node) => node.type === "process")
    .map((node) => node.id);

  const lines = [`flowchart ${graph.direction || DEFAULT_LAYOUT}`, ""];

  lines.push(
    "classDef process fill:#FFFFFF,stroke:#12345B,stroke-width:1.5px,color:#243042;"
  );
  lines.push(
    "classDef decision fill:#FFFFFF,stroke:#12345B,stroke-width:1.5px,color:#12345B;"
  );
  lines.push(
    "classDef terminal fill:#FFFFFF,stroke:#12345B,stroke-width:1.5px,color:#12345B;"
  );
  lines.push("");

  graph.nodes.forEach((node) => {
    lines.push(`    ${renderNodeCode(node)}`);
  });

  lines.push("");

  graph.edges.forEach((edge) => {
    const label = edge.label ? `|${sanitizeEdgeLabel(edge.label)}|` : "";
    lines.push(`    ${edge.source} -->${label} ${edge.target}`);
  });

  lines.push("");

  if (processIds.length) {
    lines.push(`    class ${processIds.join(",")} process`);
  }

  if (decisionIds.length) {
    lines.push(`    class ${decisionIds.join(",")} decision`);
  }

  if (terminalIds.length) {
    lines.push(`    class ${terminalIds.join(",")} terminal`);
  }

  return lines.join("\n");
}

function renderNodeCode(node) {
  const label = sanitizeCodeLabel(node.label);
  if (node.type === "decision") {
    return `${node.id}{${label}}`;
  }

  if (node.type === "start" || node.type === "end") {
    return `${node.id}([${label}])`;
  }

  return `${node.id}[${label}]`;
}

function sanitizeCodeLabel(label) {
  return cleanLabel(label)
    .replace(/\|/g, "/")
    .replace(/\[/g, "(")
    .replace(/\]/g, ")")
    .replace(/\{/g, "(")
    .replace(/\}/g, ")");
}

function sanitizeEdgeLabel(label) {
  return cleanLabel(label).replace(/\|/g, "/");
}

function setCodeEditorValue(code) {
  suppressCodeSync = true;
  refs.codeEditor.value = code;
  suppressCodeSync = false;
}

function copyCodeToClipboard() {
  const code = refs.codeEditor.value;
  if (!code.trim()) {
    return;
  }

  const fallbackCopy = () => {
    refs.codeEditor.focus();
    refs.codeEditor.select();
    document.execCommand("copy");
  };

  if (navigator.clipboard?.writeText) {
    navigator.clipboard
      .writeText(code)
      .then(() => setCodeStatus("Code copied to the clipboard.", "success"))
      .catch(() => {
        fallbackCopy();
        setCodeStatus("Code copied using the browser fallback.", "success");
      });
    return;
  }

  fallbackCopy();
  setCodeStatus("Code copied using the browser fallback.", "success");
}

async function downloadDiagram() {
  if (!cy || cy.elements().length === 0) {
    setDiagramStatus("Create a diagram before downloading it.", "No diagram");
    return;
  }

  try {
    await waitForExportFonts();
    const title = inferDownloadTitle();
    const timestamp = formatDownloadTimestamp(new Date());
    const dataUrl = cy.png({
      full: true,
      scale: 3,
      bg: "#ffffff",
    });
    const exportCanvas = await composeDiagramExport({
      title,
      timestamp,
      diagramDataUrl: dataUrl,
    });
    await downloadCanvasImage(exportCanvas, `${buildDownloadName(title)}.png`);
    setDiagramStatus("Diagram downloaded as a polished PNG.", "Downloaded");
  } catch (error) {
    console.error("Unable to export diagram.", error);
    setDiagramStatus("The diagram could not be downloaded right now.", "Export issue");
  }
}

async function composeDiagramExport({ title, timestamp, diagramDataUrl }) {
  const diagramImage = await loadDataUrlImage(diagramDataUrl);
  const framePadding = 28;
  const outerPaddingX = 120;
  const outerPaddingTop = 96;
  const outerPaddingBottom = 112;
  const minCanvasWidth = 1440;
  const maxCanvasWidth = 2800;
  const titleFontSize = 42;
  const timestampFontSize = 22;
  const titleLineHeight = 50;
  const timestampLineHeight = 30;
  const titleGap = 12;
  const dividerGap = 22;
  const sectionGap = 42;
  const diagramMaxWidth = clampNumber(
    diagramImage.naturalWidth,
    980,
    maxCanvasWidth - outerPaddingX * 2 - framePadding * 2
  );
  const diagramScale = Math.min(
    diagramMaxWidth / Math.max(diagramImage.naturalWidth, 1),
    1.08
  );
  const diagramWidth = Math.round(diagramImage.naturalWidth * diagramScale);
  const diagramHeight = Math.round(diagramImage.naturalHeight * diagramScale);
  const frameWidth = diagramWidth + framePadding * 2;
  const frameHeight = diagramHeight + framePadding * 2;
  const canvasWidth = clampNumber(
    Math.max(frameWidth + outerPaddingX * 2, minCanvasWidth),
    minCanvasWidth,
    maxCanvasWidth
  );
  const measureCanvas = document.createElement("canvas");
  const measureContext = measureCanvas.getContext("2d");
  measureContext.font = `700 ${titleFontSize}px "Sora", "Public Sans", "Segoe UI", sans-serif`;
  const titleLines = wrapCanvasText(measureContext, title, canvasWidth - outerPaddingX * 2, 2);
  const titleHeight = titleLines.length * titleLineHeight;
  const canvasHeight =
    outerPaddingTop +
    titleHeight +
    titleGap +
    timestampLineHeight +
    dividerGap +
    sectionGap +
    frameHeight +
    outerPaddingBottom;
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  canvas.width = canvasWidth;
  canvas.height = canvasHeight;
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";

  context.fillStyle = "#FBFCFE";
  context.fillRect(0, 0, canvas.width, canvas.height);

  const centerX = canvas.width / 2;
  let cursorY = outerPaddingTop;

  context.textAlign = "center";
  context.textBaseline = "top";
  context.fillStyle = "#12345B";
  context.font = `700 ${titleFontSize}px "Sora", "Public Sans", "Segoe UI", sans-serif`;
  titleLines.forEach((line) => {
    context.fillText(line, centerX, cursorY);
    cursorY += titleLineHeight;
  });

  cursorY += titleGap;
  context.fillStyle = "#6E7C8D";
  context.font = `500 ${timestampFontSize}px "Public Sans", "Segoe UI", sans-serif`;
  context.fillText(timestamp, centerX, cursorY);
  cursorY += timestampLineHeight;

  const dividerY = cursorY + dividerGap;
  context.strokeStyle = "rgba(18, 52, 91, 0.12)";
  context.lineWidth = 2;
  context.lineCap = "round";
  context.beginPath();
  context.moveTo(centerX - 84, dividerY);
  context.lineTo(centerX + 84, dividerY);
  context.stroke();

  const frameX = (canvas.width - frameWidth) / 2;
  const frameY = dividerY + sectionGap;
  drawRoundedRect(context, frameX, frameY, frameWidth, frameHeight, 28);
  context.fillStyle = "#FFFFFF";
  context.shadowColor = "rgba(13, 39, 70, 0.08)";
  context.shadowBlur = 24;
  context.shadowOffsetY = 10;
  context.fill();
  context.shadowColor = "transparent";
  context.strokeStyle = "rgba(18, 52, 91, 0.08)";
  context.lineWidth = 1.5;
  context.stroke();

  context.drawImage(
    diagramImage,
    frameX + framePadding,
    frameY + framePadding,
    diagramWidth,
    diagramHeight
  );

  return canvas;
}

function loadDataUrlImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Unable to read the rendered diagram image."));
    image.src = dataUrl;
  });
}

function downloadCanvasImage(canvas, fileName) {
  return new Promise((resolve, reject) => {
    const triggerDownload = (href) => {
      const anchor = document.createElement("a");
      anchor.href = href;
      anchor.download = fileName;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
    };

    if (typeof canvas.toBlob === "function") {
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error("Unable to prepare the PNG export."));
          return;
        }

        const url = URL.createObjectURL(blob);
        triggerDownload(url);
        window.setTimeout(() => URL.revokeObjectURL(url), 1000);
        resolve();
      }, "image/png");
      return;
    }

    try {
      triggerDownload(canvas.toDataURL("image/png"));
      resolve();
    } catch (error) {
      reject(error);
    }
  });
}

async function waitForExportFonts() {
  if (!document.fonts?.ready) {
    return;
  }

  try {
    await document.fonts.ready;
  } catch (error) {
    console.warn("Export fonts were not fully ready before rendering.", error);
  }
}

function inferDownloadTitle() {
  const prompt = cleanLabel(state.prompt);
  const theme = detectPromptTheme(prompt);
  if (theme) {
    return theme;
  }

  const firstLine = prompt
    .split(/\r?\n/)
    .map((line) => trimSentence(line))
    .find(Boolean);

  if (firstLine) {
    return buildTitleFromPromptLine(firstLine);
  }

  const firstProcessNode = state.graph?.nodes.find((node) => node.type === "process");
  if (firstProcessNode) {
    return buildTitleFromPromptLine(firstProcessNode.label);
  }

  return "Workflow Diagram";
}

function detectPromptTheme(prompt) {
  const normalizedPrompt = normalize(prompt);
  const themes = [
    { pattern: /\bexpense\b|\bfinance\b|\breimbursement\b/, title: "Expense Approval Flow" },
    { pattern: /\bsupport\b|\bticket\b|\bissue\b|\bescalat/, title: "Support Ticket Process" },
    { pattern: /\bcontent\b|\barticle\b|\beditor\b|\bwriter\b|\bpublish/, title: "Content Publishing Workflow" },
    { pattern: /\bstudent\b|\badmission\b|\badmissions\b|\bapplication\b|\bdocuments\b/, title: "Student Application Process" },
    { pattern: /\bchange management\b|\bchange request\b/, title: "Change Management Workflow" },
    { pattern: /\bproject request\b|\bproject owner\b|\bexecution\b|\bdeliverables\b/, title: "Project Request Workflow" },
    { pattern: /\bapproval\b|\bapproved\b|\brejected\b/, title: "Approval Workflow" },
    { pattern: /\bapplication\b|\bapplicant\b/, title: "Application Review Process" },
    { pattern: /\brequest\b/, title: "Request Workflow" },
  ];

  return themes.find((entry) => entry.pattern.test(normalizedPrompt))?.title || "";
}

function buildTitleFromPromptLine(line) {
  const cleaned = trimSentence(line)
    .replace(/^the flowchart should\b/i, "")
    .replace(/^flowchart\b/i, "")
    .replace(/^diagram\b/i, "")
    .replace(/^make the flowchart\b/i, "")
    .trim();
  const theme = detectPromptTheme(cleaned);
  if (theme) {
    return theme;
  }

  const condensed = cleaned
    .replace(/^(?:a|an|the)\s+/i, "")
    .replace(/\b(?:should|must|can|will)\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();
  const words = condensed.split(" ").filter(Boolean).slice(0, 5);
  const phrase = toTitleCase(words.join(" "));
  if (!phrase) {
    return "Workflow Diagram";
  }

  return /\b(flow|workflow|process|diagram)\b/i.test(phrase)
    ? phrase
    : `${phrase} Workflow`;
}

function formatDownloadTimestamp(date) {
  const formatter = new Intl.DateTimeFormat("en-US", {
    dateStyle: "long",
    timeStyle: "short",
  });
  return `Downloaded on: ${formatter.format(date)}`;
}

function wrapCanvasText(context, text, maxWidth, maxLines) {
  const words = cleanLabel(text).split(" ").filter(Boolean);
  if (!words.length) {
    return ["Workflow Diagram"];
  }

  const lines = [];
  let currentLine = words.shift();

  words.forEach((word) => {
    const nextLine = `${currentLine} ${word}`;
    if (context.measureText(nextLine).width <= maxWidth) {
      currentLine = nextLine;
      return;
    }

    lines.push(currentLine);
    currentLine = word;
  });

  lines.push(currentLine);

  if (lines.length <= maxLines) {
    return lines;
  }

  const trimmed = lines.slice(0, maxLines);
  while (
    trimmed[maxLines - 1].length > 1 &&
    context.measureText(`${trimmed[maxLines - 1]}...`).width > maxWidth
  ) {
    trimmed[maxLines - 1] = trimmed[maxLines - 1].slice(0, -1).trim();
  }
  trimmed[maxLines - 1] = `${trimmed[maxLines - 1]}...`;
  return trimmed;
}

function drawRoundedRect(context, x, y, width, height, radius) {
  const safeRadius = Math.min(radius, width / 2, height / 2);
  context.beginPath();
  context.moveTo(x + safeRadius, y);
  context.lineTo(x + width - safeRadius, y);
  context.quadraticCurveTo(x + width, y, x + width, y + safeRadius);
  context.lineTo(x + width, y + height - safeRadius);
  context.quadraticCurveTo(x + width, y + height, x + width - safeRadius, y + height);
  context.lineTo(x + safeRadius, y + height);
  context.quadraticCurveTo(x, y + height, x, y + height - safeRadius);
  context.lineTo(x, y + safeRadius);
  context.quadraticCurveTo(x, y, x + safeRadius, y);
  context.closePath();
}

function clampNumber(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function toTitleCase(text) {
  return String(text)
    .toLowerCase()
    .replace(/\b\w/g, (character) => character.toUpperCase())
    .trim();
}

function setCodeStatus(message, tone = "") {
  refs.codeStatus.textContent = message;
  refs.codeStatus.classList.remove("is-error", "is-success");

  if (tone === "error") {
    refs.codeStatus.classList.add("is-error");
  } else if (tone === "success") {
    refs.codeStatus.classList.add("is-success");
  }
}

function setDiagramStatus(message) {
  refs.diagramStatus.textContent = message;
}

function findNode(nodeId, graph = state.graph) {
  return graph?.nodes.find((node) => node.id === nodeId) || null;
}

function findEdge(edgeId, graph = state.graph) {
  return graph?.edges.find((edge) => edge.id === edgeId) || null;
}

function nextNodeId(graph, prefix = "step") {
  let index = 1;
  let candidate = `${prefix}_${index}`;
  const ids = new Set(graph.nodes.map((node) => node.id));

  while (ids.has(candidate)) {
    index += 1;
    candidate = `${prefix}_${index}`;
  }

  return candidate;
}

function nextEdgeId(graph) {
  let index = graph.edges.length + 1;
  let candidate = `edge_${index}`;
  const ids = new Set(graph.edges.map((edge) => edge.id));

  while (ids.has(candidate)) {
    index += 1;
    candidate = `edge_${index}`;
  }

  return candidate;
}

function normalizeGraph(graph) {
  const nodes = [];
  const edges = [];
  const seenNodes = new Set();
  const seenEdges = new Set();

  graph.nodes.forEach((node, index) => {
    if (!node || !node.id || seenNodes.has(node.id)) {
      return;
    }

    seenNodes.add(node.id);
    nodes.push({
      id: node.id,
      label: cleanLabel(node.label) || humanizeId(node.id),
      type: normalizeNodeType(node.type),
      order: index,
    });
  });

  graph.edges.forEach((edge, index) => {
    if (!edge || !edge.source || !edge.target) {
      return;
    }

    const key = `${edge.source}|${edge.target}|${edge.label || ""}`;
    if (seenEdges.has(key)) {
      return;
    }

    seenEdges.add(key);
    edges.push({
      id: edge.id || `edge_${index + 1}`,
      source: edge.source,
      target: edge.target,
      label: cleanLabel(edge.label),
    });
  });

  return {
    direction: graph.direction === "LR" ? "LR" : "TB",
    nodes,
    edges,
  };
}

function normalizeNodeType(type) {
  if (type === "start" || type === "process" || type === "decision" || type === "end") {
    return type;
  }

  return "process";
}

function cloneGraph(graph) {
  return {
    direction: graph.direction,
    nodes: graph.nodes.map((node) => ({ ...node })),
    edges: graph.edges.map((edge) => ({ ...edge })),
  };
}

function humanizeId(id) {
  return String(id)
    .replace(/[_-]+/g, " ")
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function cleanLabel(label) {
  return String(label || "")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHtmlAttribute(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function wrapLabelForCanvas(label, maxWidth, context) {
  const normalized = cleanLabel(label);
  if (!normalized) {
    return [""];
  }

  const words = normalized.split(/\s+/);
  const lines = [];
  let current = "";

  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (context.measureText(next).width <= maxWidth || !current) {
      current = next;
      while (context.measureText(current).width > maxWidth) {
        const sliceIndex = findWordBreakIndex(current, maxWidth, context);
        lines.push(current.slice(0, sliceIndex));
        current = current.slice(sliceIndex);
      }
      return;
    }

    lines.push(current);
    current = word;

    while (context.measureText(current).width > maxWidth) {
      const sliceIndex = findWordBreakIndex(current, maxWidth, context);
      lines.push(current.slice(0, sliceIndex));
      current = current.slice(sliceIndex);
    }
  });

  if (current) {
    lines.push(current);
  }

  return lines;
}

function getNodeTypography(label, type) {
  const length = cleanLabel(label).length;

  if (type === "decision") {
    const fontSize = length > 54 ? 14 : 15;
    return {
      fontSize,
      lineHeight: Math.round(fontSize * 1.3),
      minTextWidth: 104,
      maxTextWidth: length > 54 ? 164 : 176,
      paddingX: 32,
      paddingY: 24,
      minContentHeight: 42,
    };
  }

  if (type === "start" || type === "end") {
    const fontSize = length > 12 ? 14 : 15;
    return {
      fontSize,
      lineHeight: Math.round(fontSize * 1.25),
      minTextWidth: 56,
      maxTextWidth: 108,
      paddingX: 22,
      paddingY: 18,
      minContentHeight: 20,
    };
  }

  const fontSize = length > 78 ? 14 : 15;
  return {
    fontSize,
    lineHeight: Math.round(fontSize * 1.32),
    minTextWidth: 126,
    maxTextWidth: length > 64 ? 220 : 204,
    paddingX: 26,
    paddingY: 22,
    minContentHeight: 24,
  };
}

function getNodeLabelMeasureContext(fontSize) {
  if (!nodeLabelMeasureContext) {
    const canvas = document.createElement("canvas");
    nodeLabelMeasureContext = canvas.getContext("2d");
  }

  nodeLabelMeasureContext.font = `600 ${fontSize}px "Public Sans", "Segoe UI", sans-serif`;
  return nodeLabelMeasureContext;
}

function findWordBreakIndex(word, maxWidth, context) {
  for (let index = word.length - 1; index > 1; index -= 1) {
    if (context.measureText(word.slice(0, index)).width <= maxWidth) {
      return index;
    }
  }

  return Math.max(1, Math.min(2, word.length));
}

function preprocessPromptText(text) {
  return normalizeExtractedText(text);
}

function expandStructuredWorkflowLine(line) {
  const cleaned = cleanWorkflowDirectiveLine(line);
  if (!cleaned) {
    return [];
  }

  const decisionSplitMatch = cleaned.match(
    /^(.*?)(?:\bto\b\s+)?(?:determine|decide|check|evaluate|verify|confirm)\s+whether\s+(.+)$/i
  );

  if (decisionSplitMatch && shouldSplitDecisionPrefix(decisionSplitMatch[1])) {
    return [
      cleanNodeLabel(decisionSplitMatch[1]),
      `Determine whether ${decisionSplitMatch[2].trim()}`,
    ];
  }

  return [cleaned];
}

function sanitizeStatementCandidate(line) {
  const cleaned = cleanWorkflowDirectiveLine(line);
  if (!cleaned) {
    return "";
  }

  if (/^(yes|no|otherwise|else)$/i.test(cleaned)) {
    return "";
  }

  return cleaned;
}

function cleanWorkflowDirectiveLine(line) {
  let cleaned = String(line || "")
    .replace(/^[\s>*•\-]+/, "")
    .replace(/^\d+[\).\]]\s*/, "")
    .replace(/^step\s+\d+[:.\-]?\s*/i, "")
    .replace(/^phase\s+\d+[:.\-]?\s*/i, "")
    .trim();

  const decisionDirectiveMatch = cleaned.match(
    /^(?:after that,\s*)?(?:include|add)\s+(?:a\s+)?decision point(?:\s+to)?\s+determine whether\s+(.+)$/i
  );
  if (decisionDirectiveMatch) {
    return `Determine whether ${decisionDirectiveMatch[1].trim()}`;
  }

  const reviewDirectiveMatch = cleaned.match(
    /^show an? (.+?) step to determine whether\s+(.+)$/i
  );
  if (reviewDirectiveMatch) {
    return `${cleanNodeLabel(reviewDirectiveMatch[1])}. Determine whether ${reviewDirectiveMatch[2].trim()}`;
  }

  cleaned = cleaned
    .replace(/^the flowchart should begin when\b/i, "When")
    .replace(/^the flowchart should start when\b/i, "When")
    .replace(/^the flowchart should\b/i, "")
    .replace(/^please\b\s*/i, "")
    .trim();

  return isDiagramDirective(cleaned) ? "" : cleaned;
}

function shouldSplitDecisionPrefix(prefix) {
  const cleanedPrefix = trimSentence(prefix).toLowerCase();
  if (!cleanedPrefix) {
    return false;
  }

  return !/\b(decision point|determine|decide|check|evaluate|verify|confirm)\b/.test(
    cleanedPrefix
  );
}

function isDiagramDirective(line) {
  const normalizedLine = normalize(line);
  return /\b(make the flowchart|make the diagram|easy to understand|clear, professional|standard process boxes|decision diamonds|clean, crisp|presentation-ready|overall look|visual style|white background)\b/.test(
    normalizedLine
  );
}

function parseArrowFlow(text) {
  const graph = createGraphBuilder();
  const nodeLookup = new Map();

  text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .forEach((line) => {
      if (!line || line.startsWith("#") || line.startsWith("//")) {
        return;
      }

      const labeled = line.match(/^(.*?)\s*(?:--?>|->)\s*\[(.+?)\]\s*(.*?)$/);
      const standard = line.match(/^(.*?)\s*(?:--?>|->)\s*(.*?)$/);
      const match = labeled || standard;
      if (!match) {
        return;
      }

      const sourceText = match[1].trim();
      const label = labeled ? cleanLabel(match[2]) : "";
      const targetText = (labeled ? match[3] : match[2]).trim();
      const sourceId = ensureNode(sourceText);
      const targetId = ensureNode(targetText);

      graph.addEdge(sourceId, targetId, label);
    });

  ensureEndpoints(graph);
  return normalizeGraph({
    direction: state.layout || DEFAULT_LAYOUT,
    nodes: graph.nodes,
    edges: graph.edges,
  });

  function ensureNode(text) {
    const key = normalize(text);
    if (nodeLookup.has(key)) {
      return nodeLookup.get(key);
    }

    const id = slugify(text) || nextNodeId(graph, "step");
    const uniqueId = ensureUniqueNodeId(graph, id);
    graph.addNode(cleanNodeLabel(text), inferNodeType(text), uniqueId);
    nodeLookup.set(key, uniqueId);
    return uniqueId;
  }
}

function ensureUniqueNodeId(graph, candidate) {
  let id = candidate;
  let index = 2;
  const ids = new Set(graph.nodes.map((node) => node.id));

  while (ids.has(id)) {
    id = `${candidate}_${index}`;
    index += 1;
  }

  return id;
}

function parseNarrativeFlow(text) {
  const statements = splitIntoStatements(text);
  const graph = createGraphBuilder();
  const context = {
    previousTerminals: [],
    activeDecision: null,
    branchAnchors: new Map(),
    firstProcessId: "",
    intakeId: "",
    reviewId: "",
  };

  const startId = graph.addNode("Start", "start", "start");
  const endId = graph.addNode("End", "end", "end");
  context.previousTerminals = [startId];

  statements.forEach((statement, index) => {
    interpretStatement(statement, index, graph, context);
  });

  addLoopEdges(graph, context);
  connectLooseEnds(graph, endId, startId);

  return normalizeGraph({
    direction: state.layout || DEFAULT_LAYOUT,
    nodes: graph.nodes,
    edges: graph.edges,
  });
}

function splitIntoStatements(text) {
  return text
    .replace(/\r/g, "")
    .split(/\n+/)
    .flatMap(expandStructuredWorkflowLine)
    .flatMap((line) => line.split(/(?<=[.!?;])\s+(?=[A-Z0-9])/))
    .map(sanitizeStatementCandidate)
    .filter(Boolean);
}

function interpretStatement(statement, index, graph, context) {
  const parsed = classifyStatement(statement, index);

  switch (parsed.kind) {
    case "trigger":
      addTriggerSequence(parsed, graph, context);
      return;
    case "decision-setup":
      addDecisionSetup(parsed, graph, context);
      return;
    case "conditional":
      addConditionalBranch(parsed, graph, context);
      return;
    case "fallback":
      addFallbackBranch(parsed, graph, context);
      return;
    case "continuation":
      if (!continueBranch(parsed, graph, context)) {
        addLinearStep(parsed.action, graph, context);
      }
      return;
    default:
      addLinearStep(parsed.action || statement, graph, context);
  }
}

function classifyStatement(statement, index) {
  const cleaned = trimSentence(statement);
  const fallbackMatch = cleaned.match(/^(otherwise|else)\b[:,]?\s*(.+)$/i);
  if (fallbackMatch) {
    return {
      kind: "fallback",
      action: fallbackMatch[2].trim(),
    };
  }

  const decisionSetupMatch = cleaned.match(
    /^(?:determine|decide|check|evaluate|verify|confirm)\s+whether\s+(.+)$/i
  );
  if (decisionSetupMatch) {
    return {
      kind: "decision-setup",
      condition: decisionSetupMatch[1].trim(),
      conditionInfo: analyzeCondition(decisionSetupMatch[1]),
    };
  }

  const conditionalMatch = cleaned.match(
    /^(if|when|unless)\s+(.+?)(?:,\s*|\s+then\s+)(.+)$/i
  );
  if (conditionalMatch) {
    const condition = conditionalMatch[2].trim();
    const action = conditionalMatch[3].trim();
    const conditionInfo = analyzeCondition(condition, conditionalMatch[1].toLowerCase());

    if (
      conditionalMatch[1].toLowerCase() === "when" &&
      index === 0 &&
      looksLikeTrigger(condition)
    ) {
      return {
        kind: "trigger",
        trigger: condition,
        action,
      };
    }

    return {
      kind: "conditional",
      condition,
      conditionInfo,
      action,
    };
  }

  const continuationMatch = cleaned.match(/^(once|after)\s+(.+?)(?:,\s*|\s+then\s+)(.+)$/i);
  if (continuationMatch) {
    return {
      kind: "continuation",
      condition: continuationMatch[2].trim(),
      conditionInfo: analyzeCondition(continuationMatch[2]),
      action: continuationMatch[3].trim(),
    };
  }

  return {
    kind: "plain",
    action: cleaned,
  };
}

function addTriggerSequence(parsed, graph, context) {
  const triggerEnd = addActionPath(parsed.trigger, graph, context, context.previousTerminals)[0];
  context.previousTerminals = [triggerEnd];
  context.activeDecision = null;
  addLinearStep(parsed.action, graph, context);
}

function addDecisionSetup(parsed, graph, context) {
  const decision = getDecisionContext(parsed.conditionInfo, graph, context);
  decision.branchEnds = decision.branchEnds || new Map();
  context.activeDecision = decision;
  context.previousTerminals = [decision.id];
}

function addConditionalBranch(parsed, graph, context) {
  const decision = getDecisionContext(parsed.conditionInfo, graph, context);
  const endId = addActionPath(parsed.action, graph, context, [decision.id], parsed.conditionInfo.branchLabel)[0];
  decision.branchEnds.set(parsed.conditionInfo.branchKey, endId);
  context.branchAnchors.set(`${parsed.conditionInfo.groupKey}:${parsed.conditionInfo.branchKey}`, {
    endId,
    decisionId: decision.id,
  });
  context.previousTerminals = Array.from(new Set(decision.branchEnds.values()));
  context.activeDecision = decision;
}

function addFallbackBranch(parsed, graph, context) {
  if (!context.activeDecision) {
    addLinearStep(parsed.action, graph, context);
    return;
  }

  const existingBranches = new Set(context.activeDecision.branchEnds.keys());
  const branchKey = existingBranches.has("yes") ? "no" : "yes";
  const branchLabel = branchKey === "no" ? "No" : "Yes";
  const endId = addActionPath(parsed.action, graph, context, [context.activeDecision.id], branchLabel)[0];
  context.activeDecision.branchEnds.set(branchKey, endId);
  context.previousTerminals = Array.from(new Set(context.activeDecision.branchEnds.values()));
}

function continueBranch(parsed, graph, context) {
  const anchor = context.branchAnchors.get(`${parsed.conditionInfo.groupKey}:${parsed.conditionInfo.branchKey}`);
  const fallbackNodeId = findContinuationNode(graph, parsed.condition);

  if (!anchor && !fallbackNodeId) {
    return false;
  }

  const origin = anchor ? anchor.endId : fallbackNodeId;
  const endId = addActionPath(parsed.action, graph, context, [origin])[0];
  context.previousTerminals = [endId];
  context.activeDecision = null;
  return true;
}

function addLinearStep(text, graph, context) {
  const endId = addActionPath(text, graph, context, context.previousTerminals)[0];
  context.previousTerminals = [endId];
  context.activeDecision = null;
}

function addActionPath(text, graph, context, origins, edgeLabel = "") {
  const labels = splitActionSegments(text);
  const ids = [];
  let previousIds = Array.isArray(origins) ? origins : [origins];

  labels.forEach((label, index) => {
    const id = nextNodeId({ nodes: graph.nodes, edges: graph.edges }, inferNodeType(label) === "decision" ? "decision" : "step");
    graph.addNode(cleanNodeLabel(label), inferNodeType(label), id);
    ids.push(id);
    previousIds.forEach((fromId) => graph.addEdge(fromId, id, index === 0 ? edgeLabel : ""));
    registerAnchor(cleanNodeLabel(label), id, graph, context);
    previousIds = [id];
  });

  return [ids[ids.length - 1]];
}

function splitActionSegments(text) {
  return trimSentence(text)
    .replace(/\bfollowed by\b/gi, " then ")
    .replace(/\band finally\b/gi, " then ")
    .split(
      /\b(?:and then|then|after that|afterward|subsequently|next|finally)\b|,\s+(?=(?:assign|define|begin|start|move|mark|notify|send|return|close|archive|review|check|verify|escalate|upload|resolve|admit|schedule|share|implement|approve|reject|revise|resubmit|plan|track|complete|conduct|finalize)\b)/gi
    )
    .map((segment) => cleanNodeLabel(segment))
    .filter(Boolean);
}

function getDecisionContext(conditionInfo, graph, context) {
  if (context.activeDecision) {
    if (context.activeDecision.groupKey === conditionInfo.groupKey) {
      return context.activeDecision;
    }

    if (
      (context.activeDecision.groupKey === "generic" || context.activeDecision.groupKey.startsWith("custom:")) &&
      context.activeDecision.branchEnds.size < 2
    ) {
      return context.activeDecision;
    }
  }

  const id = nextNodeId({ nodes: graph.nodes, edges: graph.edges }, "decision");
  graph.addNode(conditionInfo.decisionLabel, "decision", id);
  context.previousTerminals.forEach((fromId) => graph.addEdge(fromId, id, ""));

  return {
    id,
    groupKey: conditionInfo.groupKey,
    branchEnds: new Map(),
  };
}

function analyzeCondition(condition, signal = "") {
  return analyzeConditionWithSignal(condition, signal);
}

function analyzeConditionWithSignal(condition, signal) {
  const normalized = normalize(condition);
  const groupKey = inferConditionTopic(normalized);
  const branchKey = inferConditionBranch(normalized, signal);
  const branchLabel = branchKey === "no" ? "No" : "Yes";
  const decisionLabel = buildDecisionLabel(condition, groupKey);

  return makeConditionInfo(groupKey, branchKey, branchLabel, decisionLabel);
}

function makeConditionInfo(groupKey, branchKey, branchLabel, decisionLabel) {
  return {
    groupKey,
    branchKey,
    branchLabel,
    decisionLabel,
  };
}

function inferConditionBranch(condition, signal) {
  if (signal === "unless") {
    return "no";
  }

  if (
    /\b(not|no|without|missing|incomplete|rejected|reject|declined|decline|denied|deny|failed|fail|complex|invalid|unavailable|needs changes|needs change|needs revision|does not|do not|is not|are not|was not|were not|cannot|can't|isn't|aren't|doesn't|don't)\b/.test(
      condition
    )
  ) {
    return "no";
  }

  return "yes";
}

function inferConditionTopic(condition) {
  if (/\b(approve|approved|approval|reject|rejected|declined|policy|accepted|authorized)\b/.test(condition)) {
    return "approval";
  }

  if (/\b(complete|completed|incomplete|missing|document|documents|required|ready)\b/.test(condition)) {
    return "completeness";
  }

  if (/\b(simple|complex|priority|impact|resource|resources)\b/.test(condition)) {
    return "evaluation";
  }

  if (/\b(publish|published|publishing|scheduled|schedule)\b/.test(condition)) {
    return "publication";
  }

  if (/\b(fix|fixed|resolve|resolved|resolution|issue)\b/.test(condition)) {
    return "resolution";
  }

  if (/\b(valid|eligible|verified|verification)\b/.test(condition)) {
    return "validation";
  }

  return `custom:${slugify(stripConditionMarkers(condition) || condition)}`;
}

function buildDecisionLabel(condition, groupKey) {
  if (groupKey === "approval") {
    return "Is it approved?";
  }

  if (groupKey === "completeness") {
    return "Is it complete?";
  }

  if (groupKey === "evaluation") {
    return "Does it meet the review criteria?";
  }

  if (groupKey === "publication") {
    return "Is it ready to publish?";
  }

  if (groupKey === "resolution") {
    return "Is the issue resolved?";
  }

  if (groupKey === "validation") {
    return "Is it valid?";
  }

  return toQuestion(stripConditionMarkers(condition));
}

function stripConditionMarkers(text) {
  return String(text || "")
    .replace(/\bdoes not follow\b/gi, "follows")
    .replace(/\bdo not follow\b/gi, "follow")
    .replace(/\bis not\b/gi, "is")
    .replace(/\bare not\b/gi, "are")
    .replace(/\bwas not\b/gi, "was")
    .replace(/\bwere not\b/gi, "were")
    .replace(/\bdoes not\b/gi, "does")
    .replace(/\bdo not\b/gi, "do")
    .replace(/\bwithout\b/gi, "")
    .replace(/\bmissing\b/gi, "complete")
    .replace(/\bincomplete\b/gi, "complete")
    .replace(/\b(rejected|declined|denied)\b/gi, "approved")
    .replace(/\bnot\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function findContinuationNode(graph, condition) {
  const tokens = extractKeywords(condition);
  let bestId = "";
  let bestScore = 0;

  graph.nodes
    .filter((node) => node.type === "process")
    .slice()
    .reverse()
    .forEach((node) => {
      const nodeTokens = extractKeywords(node.label);
      const score = tokens.filter((token) =>
        nodeTokens.some((nodeToken) => nodeToken.startsWith(token) || token.startsWith(nodeToken))
      ).length;

      if (score > bestScore) {
        bestScore = score;
        bestId = node.id;
      }
    });

  return bestScore > 0 ? bestId : "";
}

function extractKeywords(text) {
  const stopWords = new Set([
    "a",
    "an",
    "the",
    "is",
    "are",
    "was",
    "were",
    "be",
    "being",
    "once",
    "after",
    "when",
    "if",
    "it",
    "to",
    "and",
    "then",
  ]);

  return normalize(text)
    .split(/[^a-z0-9]+/)
    .filter((token) => token.length > 2 && !stopWords.has(token))
    .map((token) => token.replace(/(ing|ed|es|s)$/g, ""))
    .filter((token) => token.length > 2);
}

function registerAnchor(label, nodeId, graph, context) {
  const lower = label.toLowerCase();

  if (!context.firstProcessId) {
    context.firstProcessId = nodeId;
  }

  if (
    !context.intakeId &&
    /\b(submit|submits|submitted|upload|uploads|uploaded|reports|report|request|draft|application|ticket)\b/.test(
      lower
    )
  ) {
    context.intakeId = nodeId;
  }

  if (
    /\b(review|reviews|reviewed|editor|finance|support|manager|admissions|leadership|technical team)\b/.test(
      lower
    )
  ) {
    context.reviewId = nodeId;
  }
}

function addLoopEdges(graph, context) {
  graph.nodes.forEach((node) => {
    const label = node.label.toLowerCase();

    if (
      /\b(correct|correction|revise|revision|resubmit|returned|return to|retry|rework|update and resubmit)\b/.test(
        label
      ) &&
      context.intakeId
    ) {
      graph.addEdge(node.id, context.intakeId, "Resubmit");
    }

    if (
      /\b(upload additional documents|additional documents|technical team|send back for review|review again)\b/.test(
        label
      ) &&
      (context.reviewId || context.intakeId)
    ) {
      graph.addEdge(node.id, context.reviewId || context.intakeId, "Review again");
    }
  });
}

function connectLooseEnds(graph, endId, startId) {
  const outgoing = new Map(graph.nodes.map((node) => [node.id, 0]));
  graph.edges.forEach((edge) => outgoing.set(edge.source, (outgoing.get(edge.source) || 0) + 1));

  graph.nodes.forEach((node) => {
    if (node.id === endId || node.id === startId) {
      return;
    }

    if ((outgoing.get(node.id) || 0) === 0) {
      graph.addEdge(node.id, endId, "");
    }
  });
}

function ensureEndpoints(graph) {
  const hasStart = graph.nodes.some((node) => node.type === "start");
  const hasEnd = graph.nodes.some((node) => node.type === "end");

  if (!hasStart) {
    graph.addNode("Start", "start", "start");
    const incoming = new Set(graph.edges.map((edge) => edge.target));
    graph.nodes
      .filter((node) => node.id !== "start" && !incoming.has(node.id))
      .forEach((node) => graph.addEdge("start", node.id, ""));
  }

  if (!hasEnd) {
    graph.addNode("End", "end", "end");
    const outgoing = new Set(graph.edges.map((edge) => edge.source));
    graph.nodes
      .filter((node) => node.id !== "end" && !outgoing.has(node.id))
      .forEach((node) => graph.addEdge(node.id, "end", ""));
  }
}

function createGraphBuilder() {
  const nodes = [];
  const edges = [];
  const edgeKeys = new Set();

  return {
    nodes,
    edges,
    addNode(label, type = "process", id = `step_${nodes.length + 1}`) {
      if (nodes.some((node) => node.id === id)) {
        return id;
      }

      nodes.push({
        id,
        label: cleanLabel(label),
        type,
      });
      return id;
    },
    addEdge(source, target, label = "") {
      const key = `${source}|${target}|${label}`;
      if (!source || !target || edgeKeys.has(key)) {
        return;
      }

      edgeKeys.add(key);
      edges.push({
        id: `edge_${edges.length + 1}`,
        source,
        target,
        label: cleanLabel(label),
      });
    },
  };
}

function upsertNode(nodes, nodesById, token) {
  const existing = nodesById.get(token.id);
  if (existing) {
    if (token.label && !existing.label) {
      existing.label = token.label;
    }
    if (token.type && existing.type === "process") {
      existing.type = token.type;
    }
    return existing;
  }

  const node = {
    id: token.id,
    label: cleanLabel(token.label) || humanizeId(token.id),
    type: normalizeNodeType(token.type),
  };
  nodes.push(node);
  nodesById.set(node.id, node);
  return node;
}

function cleanNodeLabel(text) {
  const sentence = trimSentence(text)
    .replace(/^(the|a|an)\s+/i, (match) => match.toLowerCase())
    .replace(/\bit\b/g, "it")
    .trim();

  return sentence ? sentence.charAt(0).toUpperCase() + sentence.slice(1) : "Untitled step";
}

function inferNodeType(text) {
  const lower = normalize(text);
  if (/\b(start|begin)\b/.test(lower)) {
    return "start";
  }

  if (/^(end|finish|stop)$/.test(lower) || /\bworkflow ends\b/.test(lower)) {
    return "end";
  }

  if (/\?$/.test(String(text).trim())) {
    return "decision";
  }

  if (/\b(determine|decide|check|evaluate|verify|confirm)\s+whether\b/.test(lower)) {
    return "decision";
  }

  return "process";
}

function inferTerminalType(label) {
  return /\b(end|finish|close|closed)\b/i.test(label) ? "end" : "start";
}

function looksLikeTrigger(condition) {
  return /\b(submitted|reports|received|request|draft|application|ticket)\b/.test(
    normalize(condition)
  );
}

function toQuestion(text) {
  const cleaned = trimSentence(text).replace(/\.$/, "");
  return cleaned.endsWith("?") ? cleaned : `${cleanNodeLabel(cleaned)}?`;
}

function trimSentence(text) {
  return String(text).replace(/\s+/g, " ").replace(/[.]+$/g, "").trim();
}

function normalize(text) {
  return String(text).trim().toLowerCase().replace(/\s+/g, " ");
}

function slugify(text) {
  return normalize(text)
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function buildDownloadName(title) {
  const exportTitle = cleanLabel(title) || inferDownloadTitle();
  return slugify(exportTitle || state.activeSample || "flowchart_diagram") || "flowchart_diagram";
}
