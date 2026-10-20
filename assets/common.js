// Formatos y paleta compartidos por la presentación (index.html) y el reporte PDF.
// Los cálculos viven en data/build.py; aquí solo se da formato.
(function (g) {
  const abs = Math.abs;
  const F = {
    // Titulares: $4.80M · $709K
    M: v => (v < 0 ? '−' : '') + '$' + (abs(v) / 1e6).toFixed(2) + 'M',
    K: v => (v < 0 ? '−' : '') + '$' + Math.round(abs(v) / 1e3) + 'K',
    short: v => abs(v) >= 1e6 ? F.M(v) : F.K(v),
    signShort: v => (v >= 0 ? '+' : '−') + F.short(abs(v)),
    // Tablas: $4,799,584
    mxn: v => (v < 0 ? '−' : '') + '$' + Math.round(abs(v)).toLocaleString('en-US'),
    signMxn: v => (v >= 0 ? '+' : '−') + F.mxn(abs(v)),
    int: v => Math.round(v).toLocaleString('en-US'),
    n1: v => (Math.round(v * 10) / 10).toLocaleString('en-US'),
    pct: (v, d = 1) => (v * 100).toFixed(d) + '%',
    pct0: v => Math.round(v * 100) + '%',
    signPct: (v, d = 1) => (v >= 0 ? '+' : '−') + (abs(v) * 100).toFixed(d) + '%',
    x: v => v.toFixed(1) + '×',
    fecha: iso => {
      if (!iso) return '—';
      const [y, m, d] = iso.split('-');
      return `${+d}-${['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'][m - 1]}-${y.slice(2)}`;
    },
    mes: ym => ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][+ym.slice(5, 7) - 1],
    mesLargo: ym => ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][+ym.slice(5, 7) - 1],
  };
  // Paleta de gráficas (tema original del deck: familia azul/dorado)
  const P = {
    navy: '#1F3A68', steel: '#2E5C9A', soft: '#7B96C0', proj: '#B8C5D6', gold: '#C49B3A', goldBright: '#F5A623',
    success: '#2D6A4F', danger: '#9B2C2C', ink: '#0A1929', grid: '#EEF1F4', muted: '#7A8694',
    y2025: '#C49B3A', y2026: '#1F3A68',
    seq: ['#1F3A68', '#2E5C9A', '#7B96C0', '#C49B3A', '#2D6A4F', '#B8C5D6', '#9B2C2C', '#5B6B7F'],
  };
  g.F = F; g.P = P;
})(window);
