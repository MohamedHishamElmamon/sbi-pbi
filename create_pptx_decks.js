// Creates two PPTX decks:
// 1) docs/Technical_Implementation.pptx
// 2) docs/Business_KPIs.pptx
//
// Run:
//   node create_pptx_decks.js

const path = require('path');
const PptxGen = require('pptxgenjs');
const { safeOuterShadow } = require('/home/oai/share/slides/pptxgenjs_helpers/util');

// --- Paths (screenshots provided by user) ---
const IMG_YTD = path.resolve(__dirname, 'images', 'YTD.png');
const IMG_DEFINED = path.resolve(__dirname, 'images', 'defined.png');
const IMG_CUSTOM = path.resolve(__dirname, 'images', 'custome.png');

// --- Slide constants ---
const LAYOUT = 'LAYOUT_WIDE';
// 13.333 x 7.5 in
const W = 13.333;
const H = 7.5;

const COLORS = {
  navy: '0B1B3A',
  blue: '1F6FEB',
  teal: '0EA5A8',
  gray1: '111827',
  gray2: '374151',
  gray3: '6B7280',
  gray4: 'E5E7EB',
  white: 'FFFFFF',
  bg: 'F8FAFC',
  good: '16A34A',
  bad: 'DC2626',
};

function setDeckTheme(pptx) {
  pptx.layout = LAYOUT;
  pptx.author = 'Activity KPI Dashboard';
  pptx.company = ' '; // keep blank
  pptx.subject = 'Power BI + SQL Server';
  pptx.theme = {
    headFontFace: 'Calibri',
    bodyFontFace: 'Calibri',
    lang: 'en-US',
  };
}

function addHeader(pptx, slide, title, subtitle) {
  // background
  slide.background = { color: COLORS.bg };

  // Top bar
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: W,
    h: 0.68,
    fill: { color: COLORS.white },
    line: { color: COLORS.gray4 },
  });
  slide.addText(title, {
    x: 0.6,
    y: 0.16,
    w: 9.5,
    h: 0.4,
    fontFace: 'Calibri',
    fontSize: 20,
    color: COLORS.navy,
    bold: true,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6,
      y: 0.46,
      w: 10.5,
      h: 0.22,
      fontFace: 'Calibri',
      fontSize: 11,
      color: COLORS.gray3,
    });
  }
}

function addTitleSlide(pptx, deckTitle, deckSubtitle) {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.bg };

  // hero panel
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.85,
    y: 1.45,
    w: W - 1.7,
    h: 4.15,
    fill: { color: COLORS.white },
    line: { color: COLORS.gray4 },
    radius: 14,
    shadow: safeOuterShadow('000000', 0.18, 45, 3, 2),
  });

  slide.addText(deckTitle, {
    x: 1.35,
    y: 2.05,
    w: W - 2.7,
    h: 0.9,
    fontFace: 'Calibri',
    fontSize: 40,
    bold: true,
    color: COLORS.navy,
  });
  slide.addText(deckSubtitle, {
    x: 1.35,
    y: 3.1,
    w: W - 2.7,
    h: 0.6,
    fontFace: 'Calibri',
    fontSize: 16,
    color: COLORS.gray2,
  });

  // footer
  slide.addText('As-of: Jan 20, 2026 (project context)', {
    x: 1.35,
    y: 5.15,
    w: W - 2.7,
    h: 0.3,
    fontFace: 'Calibri',
    fontSize: 11,
    color: COLORS.gray3,
  });

  slide.addNotes(
    `[Sources]\n- Date-table/time-intelligence guidance (Power BI): https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-date-tables\n` +
      `- DATESBETWEEN (DAX): https://learn.microsoft.com/en-us/dax/datesbetween-function-dax\n` +
      `[/Sources]`
  );
}

function addBullets(pptx, slide, x, y, w, title, bullets) {
  slide.addText(title, {
    x,
    y,
    w,
    h: 0.3,
    fontFace: 'Calibri',
    fontSize: 16,
    bold: true,
    color: COLORS.navy,
  });
  slide.addText(
    bullets.map((t) => ({ text: t, options: { bullet: { indent: 18 }, hanging: 6 } })),
    {
      x,
      y: y + 0.42,
      w,
      h: 2.1,
      fontFace: 'Calibri',
      fontSize: 12.5,
      color: COLORS.gray2,
      valign: 'top',
      lineSpacingMultiple: 1.15,
    }
  );
}

function addCallout(pptx, slide, x, y, w, h, title, body) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    fill: { color: COLORS.white },
    line: { color: COLORS.gray4 },
    radius: 10,
    shadow: safeOuterShadow('000000', 0.14, 45, 2.5, 2),
  });
  slide.addText(title, {
    x: x + 0.25,
    y: y + 0.18,
    w: w - 0.5,
    h: 0.3,
    fontFace: 'Calibri',
    fontSize: 14,
    bold: true,
    color: COLORS.navy,
  });
  slide.addText(body, {
    x: x + 0.25,
    y: y + 0.55,
    w: w - 0.5,
    h: h - 0.7,
    fontFace: 'Calibri',
    fontSize: 11.5,
    color: COLORS.gray2,
    valign: 'top',
  });
}

function addImageCard(pptx, slide, imgPath, x, y, w, h, caption) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    fill: { color: COLORS.white },
    line: { color: COLORS.gray4 },
    radius: 10,
    shadow: safeOuterShadow('000000', 0.18, 45, 3, 2),
  });
  slide.addImage({ path: imgPath, x: x + 0.18, y: y + 0.18, w: w - 0.36, h: h - 0.6 });
  if (caption) {
    slide.addText(caption, {
      x: x + 0.18,
      y: y + h - 0.35,
      w: w - 0.36,
      h: 0.25,
      fontFace: 'Calibri',
      fontSize: 10.5,
      color: COLORS.gray3,
    });
  }
}

function buildTechnicalDeck() {
  const pptx = new PptxGen();
  setDeckTheme(pptx);

  // Slide 1
  addTitleSlide(pptx, 'Technical Implementation', 'SQL Server (Cloud SQL) + Power BI • Activity KPI Dashboard');

  // Slide 2: Architecture
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'System overview', 'High-level data flow and modeling approach');

    // Flow boxes
    const topY = 1.25;
    const boxH = 1.05;
    const boxW = 3.55;
    const gap = 0.45;
    const x1 = 0.85;
    const x2 = x1 + boxW + gap;
    const x3 = x2 + boxW + gap;

    const boxes = [
      { x: x1, t: '1) Restore DB', b: 'Import .bak into SQL Server\n(Cloud SQL for SQL Server)', c: COLORS.blue },
      { x: x2, t: '2) Prepare dataset', b: 'Shift dates to 2025/2026\nRebuild Dim_Time_day (2025–2026)', c: COLORS.teal },
      { x: x3, t: '3) Power BI model', b: 'Star schema\nDAX measures A/B/C\n3 report pages', c: COLORS.navy },
    ];
    boxes.forEach((bx) => {
      slide.addShape(pptx.ShapeType.roundRect, {
        x: bx.x,
        y: topY,
        w: boxW,
        h: boxH,
        fill: { color: COLORS.white },
        line: { color: COLORS.gray4 },
        radius: 10,
        shadow: safeOuterShadow('000000', 0.12, 45, 2.5, 2),
      });
      slide.addShape(pptx.ShapeType.rect, {
        x: bx.x,
        y: topY,
        w: 0.08,
        h: boxH,
        fill: { color: bx.c },
        line: { color: bx.c },
      });
      slide.addText(bx.t, {
        x: bx.x + 0.2,
        y: topY + 0.12,
        w: boxW - 0.35,
        h: 0.28,
        fontFace: 'Calibri',
        fontSize: 15,
        bold: true,
        color: COLORS.navy,
      });
      slide.addText(bx.b, {
        x: bx.x + 0.2,
        y: topY + 0.45,
        w: boxW - 0.35,
        h: 0.6,
        fontFace: 'Calibri',
        fontSize: 11.5,
        color: COLORS.gray2,
      });
    });

    // Data model mini-diagram
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.85,
      y: 2.75,
      w: W - 1.7,
      h: 4.35,
      fill: { color: COLORS.white },
      line: { color: COLORS.gray4 },
      radius: 12,
    });
    slide.addText('Power BI star schema', {
      x: 1.1,
      y: 2.92,
      w: 4,
      h: 0.25,
      fontSize: 14,
      bold: true,
      color: COLORS.navy,
    });

    // Fact table
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 5.3,
      y: 3.45,
      w: 2.75,
      h: 1.05,
      fill: { color: 'EEF2FF' },
      line: { color: 'C7D2FE' },
      radius: 10,
    });
    slide.addText('v_Fact_Activities_Shifted', {
      x: 5.45,
      y: 3.58,
      w: 2.45,
      h: 0.35,
      fontSize: 12,
      bold: true,
      color: COLORS.navy,
    });
    slide.addText('• ind_id\n• activity_date\n• act_value', {
      x: 5.55,
      y: 3.92,
      w: 2.35,
      h: 0.55,
      fontSize: 11,
      color: COLORS.gray2,
    });

    // Dim tables
    const dimStyle = { fill: { color: 'ECFEFF' }, line: { color: 'A5F3FC' }, radius: 10 };
    slide.addShape(pptx.ShapeType.roundRect, { x: 2.0, y: 3.25, w: 2.4, h: 0.9, ...dimStyle });
    slide.addText('Dim_Indicator', {
      x: 2.15,
      y: 3.36,
      w: 2.1,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: COLORS.navy,
    });
    slide.addText('ind_id, ind_desc…', {
      x: 2.15,
      y: 3.63,
      w: 2.1,
      h: 0.25,
      fontSize: 10.5,
      color: COLORS.gray2,
    });

    slide.addShape(pptx.ShapeType.roundRect, { x: 2.0, y: 4.35, w: 2.4, h: 0.9, ...dimStyle });
    slide.addText('Dim_Time_day', {
      x: 2.15,
      y: 4.46,
      w: 2.1,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: COLORS.navy,
    });
    slide.addText('day_date, week_id…', {
      x: 2.15,
      y: 4.73,
      w: 2.1,
      h: 0.25,
      fontSize: 10.5,
      color: COLORS.gray2,
    });

    // Relationship arrows
    slide.addShape(pptx.ShapeType.line, { x: 4.45, y: 3.7, w: 0.85, h: 0, line: { color: COLORS.gray3, width: 2, beginArrowType: 'none', endArrowType: 'triangle' } });
    slide.addShape(pptx.ShapeType.line, { x: 4.45, y: 4.8, w: 0.85, h: -0.8, line: { color: COLORS.gray3, width: 2, beginArrowType: 'none', endArrowType: 'triangle' } });

    addBullets(pptx, slide, 1.1, 5.55, 3.9, 'Implementation notes', [
      'Date table is marked in Power BI to enable time-intelligence patterns.',
      'AsOfDate anchored to the latest FACT date (prevents empty future dates).',
      'All KPIs are filtered by selected Indicator (Dim_Indicator).',
    ]);

    slide.addNotes(
      `[Sources]\n- Mark as date table (Power BI): https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-date-tables\n` +
        `- Date table modeling guidance: https://learn.microsoft.com/en-us/power-bi/guidance/model-date-tables\n[/Sources]`
    );
  }

  // Slide 3: Restore
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Database restore (Cloud SQL for SQL Server)', 'Import .bak into the SQL Server instance');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      6.2,
      2.05,
      'Restore steps',
      '1) Upload the .bak file to Cloud Storage\n' +
        '2) Import/restore into Cloud SQL (SQL Server)\n' +
        '3) Verify database and user credentials\n' +
        '4) Confirm connectivity from Power BI (public IP or proxy)'
    );

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 7.35,
      y: 1.25,
      w: 5.15,
      h: 2.05,
      fill: { color: '0B1220' },
      line: { color: '0B1220' },
      radius: 10,
    });
    slide.addText('Example CLI', {
      x: 7.6,
      y: 1.42,
      w: 4.7,
      h: 0.3,
      fontFace: 'Calibri',
      fontSize: 12,
      bold: true,
      color: 'FFFFFF',
    });
    slide.addText(
      'gcloud sql import bak INSTANCE \\\n+  gs://BUCKET/backup.bak \\\n+  --database=DB_NAME',
      {
        x: 7.6,
        y: 1.78,
        w: 4.8,
        h: 1.4,
        fontFace: 'Consolas',
        fontSize: 12,
        color: 'D1D5DB',
        valign: 'top',
      }
    );

    addCallout(
      pptx,
      slide,
      0.85,
      3.55,
      W - 1.7,
      3.65,
      'Why this matters',
      'The restored database is the single source used by Power BI. The remaining steps (date shifting, date dimension rebuild, measures) are layered on top without changing the original backup tables.'
    );

    slide.addNotes(
      `[Sources]\n- Cloud SQL for SQL Server import/export with BAK: https://docs.cloud.google.com/sql/docs/sqlserver/import-export/import-export-bak\n` +
        `- gcloud sql import bak reference: https://docs.cloud.google.com/sdk/gcloud/reference/sql/import/bak\n[/Sources]`
    );
  }

  // Slide 4: Data preparation (SQL)
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Data preparation (SQL)', 'Shift dates to 2025/2026 and rebuild the date dimension');

    addBullets(pptx, slide, 0.85, 1.15, 6.0, 'Key changes', [
      'Create a view (v_Fact_Activities_Shifted) that moves activity_date forward by +4 years (2021→2025, 2022→2026).',
      'Cast act_value to INT (defaults invalid values to 0).',
      'Rebuild Dim_Time_day for a full daily calendar (2025-01-01 → 2026-12-31) with ISO week_id and LY/LW helper dates.',
    ]);

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 7.05,
      y: 1.15,
      w: 5.43,
      h: 5.75,
      fill: { color: '0B1220' },
      line: { color: '0B1220' },
      radius: 10,
    });
    slide.addText('Snippet (shift view)', {
      x: 7.3,
      y: 1.32,
      w: 4.95,
      h: 0.3,
      fontFace: 'Calibri',
      fontSize: 12,
      bold: true,
      color: 'FFFFFF',
    });
    slide.addText(
      "DATEADD(year, 4, CONVERT(date,\n  TRY_CONVERT(datetimeoffset(0), activity_date)))\nAS activity_date",
      {
        x: 7.3,
        y: 1.7,
        w: 5.0,
        h: 1.2,
        fontFace: 'Consolas',
        fontSize: 12,
        color: 'D1D5DB',
        valign: 'top',
      }
    );
    slide.addText('Snippet (ISO week_id)', {
      x: 7.3,
      y: 3.1,
      w: 5.0,
      h: 0.3,
      fontFace: 'Calibri',
      fontSize: 12,
      bold: true,
      color: 'FFFFFF',
    });
    slide.addText(
      "CONCAT(\n  YEAR(DATEADD(day,3, DATETRUNC(iso_week, d))),\n  RIGHT('00'+CAST(DATEPART(iso_week,d) AS varchar(2)),2)\n) AS week_id",
      {
        x: 7.3,
        y: 3.48,
        w: 5.05,
        h: 1.55,
        fontFace: 'Consolas',
        fontSize: 11.2,
        color: 'D1D5DB',
        valign: 'top',
      }
    );

    addCallout(
      pptx,
      slide,
      0.85,
      5.95,
      6.0,
      1.0,
      'Outcome',
      'Power BI can treat the dataset as “current/previous year” and compute YTD/LY and rolling windows using a complete date dimension.'
    );

    slide.addNotes(
      `[Sources]\n- Time-intelligence patterns rely on a proper date table: https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-date-tables\n[/Sources]`
    );
  }

  // Slide 5: Power BI model
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Power BI model', 'Tables loaded, relationships, and refresh');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      6.25,
      2.35,
      'Imported tables',
      '• dbo.v_Fact_Activities_Shifted\n' +
        '• dbo.Dim_Indicator\n' +
        '• dbo.Dim_Time_day\n\n' +
        'Optional: Dim_Employees for future slicing (not required for KPIs).'
    );

    addCallout(
      pptx,
      slide,
      0.85,
      3.85,
      6.25,
      3.05,
      'Relationships + settings',
      '• Fact[ind_id] → Dim_Indicator[ind_id]\n' +
        '• Fact[activity_date] → Dim_Time_day[day_date]\n' +
        '• Mark Dim_Time_day as the Date table\n' +
        '• Use AsOfDate = max fact date to avoid empty future periods'
    );

    // Screenshot montage
    slide.addText('Report pages (examples)', {
      x: 7.35,
      y: 1.25,
      w: 5.15,
      h: 0.3,
      fontSize: 14,
      bold: true,
      color: COLORS.navy,
    });
    addImageCard(pptx, slide, IMG_YTD, 7.35, 1.6, 5.15, 1.85, 'YTD');
    addImageCard(pptx, slide, IMG_DEFINED, 7.35, 3.55, 5.15, 1.85, 'Defined period (365/180)');
    addImageCard(pptx, slide, IMG_CUSTOM, 7.35, 5.5, 5.15, 1.85, 'Custom period comparison');

    slide.addNotes(
      `[Sources]\n- Mark as date table (Power BI): https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-date-tables\n` +
        `- Date table modeling guidance: https://learn.microsoft.com/en-us/power-bi/guidance/model-date-tables\n[/Sources]`
    );
  }

  // Slide 6: KPI calculations
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'KPI calculations (DAX)', 'Measures for A) YTD, B) rolling windows, C) custom periods');

    addBullets(pptx, slide, 0.85, 1.15, 6.2, 'Core measures', [
      'Activities = SUM(Fact[act_value])',
      'AsOfDate = MAX(Fact[activity_date]) with filters removed',
      'All comparisons are scoped to the selected Indicator (Dim_Indicator).',
    ]);

    addCallout(
      pptx,
      slide,
      0.85,
      3.35,
      6.2,
      3.55,
      'A/B/C logic summary',
      'A) YTD: Jan 1 → AsOfDate; LY = same window 12 months earlier\n' +
        'B) Defined: Last N days (365/180) → AsOfDate; LY = same shifted window\n' +
        'C) Custom: Period1 vs Period2 defined by slicers; trend shows both periods\n\n' +
        'Variance % = (Current − LY) / LY; arrow ▲/▼ based on sign'
    );

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 7.35,
      y: 1.15,
      w: 5.15,
      h: 5.75,
      fill: { color: '0B1220' },
      line: { color: '0B1220' },
      radius: 10,
    });
    slide.addText('Example pattern', {
      x: 7.6,
      y: 1.32,
      w: 4.7,
      h: 0.3,
      fontFace: 'Calibri',
      fontSize: 12,
      bold: true,
      color: 'FFFFFF',
    });
    slide.addText(
      'm_YTD =\nVAR d = [AsOfDate]\nVAR s = DATE(YEAR(d),1,1)\nRETURN CALCULATE([Activities],\n  DATESBETWEEN(Dim_Time_day[day_date], s, d))',
      {
        x: 7.6,
        y: 1.7,
        w: 4.85,
        h: 2.2,
        fontFace: 'Consolas',
        fontSize: 11.2,
        color: 'D1D5DB',
        valign: 'top',
      }
    );
    slide.addText(
      'm_VarPct =\nVAR cur=[Current]\nVAR ly=[LY]\nRETURN DIVIDE(cur-ly, ly)',
      {
        x: 7.6,
        y: 4.1,
        w: 4.85,
        h: 1.25,
        fontFace: 'Consolas',
        fontSize: 11.2,
        color: 'D1D5DB',
      }
    );

    slide.addNotes(
      `[Sources]\n- DATESBETWEEN (DAX): https://learn.microsoft.com/en-us/dax/datesbetween-function-dax\n` +
        `- DAX time intelligence functions: https://learn.microsoft.com/en-us/dax/time-intelligence-functions-dax\n[/Sources]`
    );
  }

  // Slide 7: Navigation (pages or bookmarks)
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Report UX', 'Three pages (YTD / Defined / Custom) or single-page tabs via bookmarks');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      6.2,
      2.2,
      'Navigation options',
      'Option 1 (simple): Three report pages\n' +
        '• YTD page\n• Defined Period page\n• Custom Period page\n\n' +
        'Option 2 (tabbed): One page + bookmark navigator\n' +
        '• Group visuals by tab (Selection pane)\n• Create bookmarks per tab\n• Use Bookmark navigator buttons'
    );

    addBullets(pptx, slide, 0.85, 3.75, 6.2, 'Slicer behavior', [
      'Keep Indicator slicer global (applies across tabs/pages).',
      'For bookmark tabs, disable “Data” for bookmarks to prevent slicer resets.',
      'Trend charts use a trend measure that returns BLANK outside the selected window.',
    ]);

    addImageCard(pptx, slide, IMG_DEFINED, 7.35, 1.25, 5.15, 5.9, 'Example: Defined period page');

    slide.addNotes(
      `[Sources]\n- Bookmarks in Power BI: https://learn.microsoft.com/en-us/power-bi/create-reports/desktop-bookmarks\n` +
        `- Page & bookmark navigators: https://learn.microsoft.com/en-us/power-bi/create-reports/button-navigators\n[/Sources]`
    );
  }

  // Slide 8: Refresh & “today” filtering
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Refresh & “today” filtering', 'Optional Power Query filter to restrict dataset to today');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      6.2,
      1.85,
      'Why filter to today?',
      'When the date dimension extends beyond available fact data, visuals may show empty future dates. Filtering to today (or using AsOfDate based on fact max) keeps KPIs aligned with the “current” reporting cut.'
    );

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.85,
      y: 3.35,
      w: 6.2,
      h: 3.85,
      fill: { color: '0B1220' },
      line: { color: '0B1220' },
      radius: 10,
    });
    slide.addText('Power Query (M) example', {
      x: 1.1,
      y: 3.52,
      w: 5.8,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: 'FFFFFF',
    });
    slide.addText(
      'TodayUTC = Date.From(DateTimeZone.UtcNow()),\n' +
        'FilteredToToday = Table.SelectRows(\n' +
        '  #"Changed Type", each [activity_date] <= TodayUTC)',
      {
        x: 1.1,
        y: 3.9,
        w: 5.9,
        h: 1.4,
        fontFace: 'Consolas',
        fontSize: 11.6,
        color: 'D1D5DB',
        valign: 'top',
      }
    );

    addCallout(
      pptx,
      slide,
      7.35,
      1.25,
      5.15,
      5.95,
      'Recommended approach used in this project',
      '• AsOfDate is based on the latest fact date, so KPIs do not depend on the end of the date table.\n' +
        '• Optional M filter ensures refresh only includes data up to today (UTC) when required.'
    );

    slide.addNotes(
      `[Sources]\n- Date table settings (Power BI): https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-date-tables\n[/Sources]`
    );
  }

  // Slide 9: QA + handoff
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'QA, validation & handoff', 'How results were checked and what is delivered');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      6.2,
      2.4,
      'Validation checks',
      '• SQL spot-check: SUM(act_value) for an indicator over a known date window\n' +
        '• Power BI cards match SQL results for the same window\n' +
        '• Relationships verified (Fact→Date, Fact→Indicator)\n' +
        '• Edge cases: LY = 0 handled with DIVIDE() to avoid errors'
    );

    addCallout(
      pptx,
      slide,
      0.85,
      4.0,
      6.2,
      2.95,
      'Handoff package (GitHub)',
      '• SQL scripts (shift view + rebuild date dimension)\n' +
        '• Power BI file (PBIX)\n' +
        '• Two decks: Technical + Business\n' +
        '• README with setup steps and troubleshooting'
    );

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 7.35,
      y: 1.25,
      w: 5.15,
      h: 5.7,
      fill: { color: COLORS.white },
      line: { color: COLORS.gray4 },
      radius: 10,
    });
    slide.addText('Deliverables', {
      x: 7.6,
      y: 1.45,
      w: 4.7,
      h: 0.3,
      fontSize: 14,
      bold: true,
      color: COLORS.navy,
    });
    slide.addText(
      [
        { text: '• Activity-KPI-Dashboard.pbix', options: { bullet: { indent: 18 }, hanging: 6 } },
        { text: '• Technical_Implementation.pptx', options: { bullet: { indent: 18 }, hanging: 6 } },
        { text: '• Business_KPIs.pptx', options: { bullet: { indent: 18 }, hanging: 6 } },
        { text: '• SQL scripts + README', options: { bullet: { indent: 18 }, hanging: 6 } },
      ],
      {
        x: 7.6,
        y: 1.85,
        w: 4.8,
        h: 1.4,
        fontSize: 12,
        color: COLORS.gray2,
      }
    );

    slide.addNotes(`[Sources]\n- DAX time-intelligence overview: https://learn.microsoft.com/en-us/dax/time-intelligence-functions-dax\n[/Sources]`);
  }

  return pptx;
}

function buildBusinessDeck() {
  const pptx = new PptxGen();
  setDeckTheme(pptx);

  // Slide 1
  addTitleSlide(pptx, 'Business KPIs & Insights', 'How to read the dashboard and what it enables for decision-making');

  // Slide 2: What the dashboard answers
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'What this dashboard answers', 'One place to track activity volume by Indicator, compare periods, and spot trend changes');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      6.1,
      2.55,
      'Primary questions',
      '• Are we up or down vs last year for the same period?\n' +
        '• Is the change driven by specific activity Indicators (email, meetings, etc.)?\n' +
        '• Are there spikes/drops that require operational follow-up?\n' +
        '• How do two business-defined periods compare (custom Period1 vs Period2)?'
    );
    addCallout(
      pptx,
      slide,
      0.85,
      4.05,
      6.1,
      3.1,
      'How to use (workflow)',
      '1) Select Indicator (single-select)\n' +
        '2) Choose the analysis mode: YTD, Defined (365/180), or Custom\n' +
        '3) Read the KPI cards (Current, LY, % change)\n' +
        '4) Use the trend chart to interpret timing and volatility'
    );

    addImageCard(pptx, slide, IMG_YTD, 7.35, 1.25, 5.15, 5.9, 'Example: YTD page');

    slide.addNotes(
      `[Sources]\n- Date table/time intelligence patterns in Power BI: https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-date-tables\n[/Sources]`
    );
  }

  // Slide 3: KPI definitions
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'KPI definitions', 'What YTD / LY / rolling windows / custom periods mean');

    addCallout(
      pptx,
      slide,
      0.85,
      1.25,
      W - 1.7,
      1.75,
      'Definitions (as-of Jan 20, 2026)',
      '• YTD: Jan 1, 2026 → AsOfDate (latest loaded date with activity)\n' +
        '• YTD LY: Jan 1, 2025 → same day-of-year cut (12 months earlier)\n' +
        '• Last 365/180: trailing window ending at AsOfDate; LY = same shifted window\n' +
        '• Custom periods: user-selected Period1 vs Period2 ranges\n' +
        '• % variance: (Current − Comparison) / Comparison\n' +
        '• Arrow: ▲ when % ≥ 0, ▼ when % < 0'
    );

    addCallout(
      pptx,
      slide,
      0.85,
      3.25,
      6.1,
      3.95,
      'Interpretation tips',
      '• Large % swings on small LY values can be noisy—use the trend line to confirm.\n' +
        '• Look for sustained changes (multi-week) vs single-day spikes.\n' +
        '• Use Defined Period mode for short-term operational monitoring.'
    );

    addCallout(
      pptx,
      slide,
      7.35,
      3.25,
      5.15,
      3.95,
      'What counts as “activity”',
      'The fact table stores an Indicator ID and an activity count/value per date. Indicator metadata (category/subcategory) enables slicing and consistent reporting across activity types.'
    );

    slide.addNotes(
      `[Sources]\n- DATESBETWEEN (DAX): https://learn.microsoft.com/en-us/dax/datesbetween-function-dax\n` +
        `- DAX time intelligence functions: https://learn.microsoft.com/en-us/dax/time-intelligence-functions-dax\n[/Sources]`
    );
  }

  // Slide 4: Defined period example
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Defined period (Last 365 / 180 days)', 'Operational monitoring with consistent windows');

    addImageCard(pptx, slide, IMG_DEFINED, 0.85, 1.25, W - 1.7, 5.95, 'Defined Period view (example screenshot)');

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.95,
      y: 6.9,
      w: W - 1.9,
      h: 0.45,
      fill: { color: COLORS.white },
      line: { color: COLORS.gray4 },
      radius: 10,
    });
    slide.addText(
      'Use this view to answer: “Are we trending up/down over the last N days, and how does it compare to last year’s same period?”',
      {
        x: 1.15,
        y: 7.0,
        w: W - 2.3,
        h: 0.28,
        fontSize: 12,
        color: COLORS.gray2,
      }
    );

    slide.addNotes(`[Sources]\n- Bookmarks & navigation (optional): https://learn.microsoft.com/en-us/power-bi/create-reports/button-navigators\n[/Sources]`);
  }

  // Slide 5: Custom periods example
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Custom periods (Period 1 vs Period 2)', 'Compare any two business-defined windows');

    addImageCard(pptx, slide, IMG_CUSTOM, 0.85, 1.25, W - 1.7, 5.95, 'Custom Period view (example screenshot)');

    addCallout(
      pptx,
      slide,
      0.85,
      6.55,
      W - 1.7,
      0.85,
      'Typical uses',
      'Compare pre/post policy changes, campaigns, org changes, holidays, or project milestones by selecting two date ranges and evaluating both value and trend.'
    );

    slide.addNotes(`[Sources]\n- Date table/time intelligence patterns: https://learn.microsoft.com/en-us/power-bi/guidance/model-date-tables\n[/Sources]`);
  }

  // Slide 6: Business impact and next steps
  {
    const slide = pptx.addSlide();
    addHeader(pptx, slide, 'Business impact & next steps', 'How this dashboard supports decision-making');

    addBullets(pptx, slide, 0.85, 1.2, 6.25, 'Business impact', [
      'Single source for activity KPIs across Indicators with consistent comparisons.',
      'Faster trend detection: spikes/drops are visible immediately in the trend charts.',
      'Supports operational planning (resource allocation, workload patterns) using rolling windows.',
      'Supports business review cycles via custom period comparison (Period1 vs Period2).',
    ]);

    addBullets(pptx, slide, 0.85, 3.95, 6.25, 'Suggested enhancements (optional)', [
      'Add drill-through to employee/department (Dim_Employees) when needed.',
      'Add indicator category rollups and a “Top movers” view.',
      'Add anomaly flags (simple z-score) to highlight outlier days.',
      'Publish to Power BI Service + scheduled refresh (gateway/proxy as required).',
    ]);

    addCallout(
      pptx,
      slide,
      7.35,
      1.2,
      5.15,
      5.7,
      'What stakeholders receive',
      '• A self-serve dashboard with three analysis modes\n' +
        '• Clear definitions of KPIs (YTD, LY, rolling, custom)\n' +
        '• Visual trend context to interpret changes\n' +
        '• A technical README + scripts for reproducibility'
    );

    slide.addNotes(
      `[Sources]\n- Power BI bookmarks (for tabbed UX): https://learn.microsoft.com/en-us/power-bi/create-reports/desktop-bookmarks\n` +
        `- Power BI navigators: https://learn.microsoft.com/en-us/power-bi/create-reports/button-navigators\n[/Sources]`
    );
  }

  return pptx;
}

// --- Main ---
const tech = buildTechnicalDeck();
const biz = buildBusinessDeck();

const outTech = path.resolve(__dirname, 'docs', 'Technical_Implementation.pptx');
const outBiz = path.resolve(__dirname, 'docs', 'Business_KPIs.pptx');

tech.writeFile({ fileName: outTech });
biz.writeFile({ fileName: outBiz });

console.log('Wrote:', outTech);
console.log('Wrote:', outBiz);
