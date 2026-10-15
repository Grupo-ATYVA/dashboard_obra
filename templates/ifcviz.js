const IFCViz = (() => {
  let ready = false;

  function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function init() {
    if (ready) return;
    ready = true;
    setText('ifcviz-conn-lbl', 'Use a aba IFC 3D para carregar e medir o modelo localmente.');
  }

  function loadFile(file) {
    init();
    setText('ifcviz-fname', file ? file.name : '');
    if (window.IFC && typeof window.IFC.loadFile === 'function') {
      window.IFC.loadFile(file);
    } else {
      setText('ifcviz-conn-lbl', 'Visualizador IFC local indisponivel.');
    }
  }

  function filterTable() {}

  function exportCSV() {
    alert('Exportacao IFC disponivel pela aba IFC 3D.');
  }

  function onModelLoaded() {}

  return { init, loadFile, filterTable, exportCSV, onModelLoaded };
})();
// Ver nota equivalente em ifc_viewer.js: "const" no escopo global de um
// script classico nao vira window.IFCViz sozinho.
window.IFCViz = IFCViz;
