import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabase";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend
} from "recharts";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell } from "docx";
import { saveAs } from "file-saver";
import html2canvas from "html2canvas";
import { ImageRun } from "docx";
import { AlignmentType } from "docx";
import logo from "./logo.png";


export default function App() {


  /* PROGRAMS */
  const [programs, setPrograms] = useState([]);
  const [program, setProgram] = useState("");
  const [newProgram, setNewProgram] = useState("");
  const [plos, setPlos] = useState([]);
  const [user, setUser] = useState(null);
    
useEffect(() => {
  (async () => {
    try {

let data = null;
let error = null;

if (supabase) {
  const res = await supabase.auth.getUser();
  data = res.data;
  error = res.error;
}

      if (error) {
        console.error(error);
        return;
      }
      setUser(data?.user || null);
    } catch (e) {
      console.error("Auth crash:", e);
    }
  })();
}, []);

const [email, setEmail] = useState("");
const [password, setPassword] = useState("");
  const [selectedPLOs, setSelectedPLOs] = useState({});
  const [editingId, setEditingId] = useState(null);
  const [editedName, setEditedName] = useState("");
  const [newPLO, setNewPLO] = useState("");
  const [editingPLO, setEditingPLO] = useState(null);
  const [showCharts, setShowCharts] = useState(true);
  const [editedPLOName, setEditedPLOName] = useState("");
  const [activeTab, setActiveTab] = useState("program");

  // COURSES
  const [courses, setCourses] = useState([]);
  const [newCourseCode, setNewCourseCode] = useState("");
  const [selectedCourseId, setSelectedCourseId] = useState("");

  /* DATA */
  const [records, setRecords] = useState([]);
  const [description, setDescription] = useState("");
  const [achievement, setAchievement] = useState("");
  const [year, setYear] = useState("");
  const [activeYear, setActiveYear] = useState("");
  // ✅ REPORT META (used by AI + PDF + Word)
  const reportMeta = {
    university: "King Saud University",
    college: "College of Architecture and Planning",
    department: "Architecture and Building Sciences",
    program,
    academicYear: activeYear,
    hijriYear: "1444–1445"
  };

  const [previousYearSelected, setPreviousYearSelected] = useState("");
  const [isCycleClosed, setIsCycleClosed] = useState(false);
  const [plo, setPlo] = useState(plos.length ? plos[0].name : "");
  const [weight, setWeight] = useState(0);
  const [newCycleYear, setNewCycleYear] = useState("");

  // FIX #5: initialise ploResults as empty object, not hardcoded PLO keys
  const [ploResults, setPloResults] = useState({});
  const [indirectResults, setIndirectResults] = useState({});

  // FIX #2: previousYear is now dynamic — looked up by PLO name at runtime
  // Stored as a ref-like state so it can be extended; default to 0 for unknown PLOs
  const [previousYear, setPreviousYear] = useState({});

  /* ─── LOAD PROGRAMS ─── */
  const loadPrograms = async () => {
    const { data } = await supabase.from("programs").select("*");
    if (data) {
      setPrograms(data);
      if (data.length) setProgram(data[0].name);
    }
  };



  useEffect(() => {
    loadPrograms();
  }, []);

  useEffect(() => {
    if (plos.length > 0) {
      setPlo(plos[0].name);
    }
  }, [plos]);

  /* ─── LOAD PLOs ─── */
  const loadPLOs = useCallback(async () => {
    if (!program) return;
    const { data } = await supabase
      .from("plos")
      .select("*")
      .eq("program", program)

    if (data) setPlos(data);
  }, [program]);

const handleLogin = async () => {
  const { data, error } = await supabase.auth.signInWithPassword({
    email,
    password
  });

  if (error) {
    alert(error.message);
  } else {
    setUser(data.user);
  }
};

const handleLogout = async () => {
  await supabase.auth.signOut();
  setUser(null);
};

  const loadCourses = useCallback(async () => {
    if (!program) return;

    const { data } = await supabase
      .from("courses")
      .select("*")
      .eq("program", program);

    if (data) setCourses(data);
  }, [program]);

  useEffect(() => {
    loadCourses();
  }, [loadCourses]);

  useEffect(() => {
    if (program) loadPLOs();
  }, [program, loadPLOs]);

  /* ─── LOAD DATA ─── */
  // FIX #3: loadData no longer calls calculatePLO directly.
  // It only sets records; the effect below triggers calculation once BOTH
  // plos and records are available, avoiding the race condition.
  const loadData = useCallback(async () => {
    if (!program) return;
    const { data } = await supabase
      .from("clo_records")
      .select("*")
      .eq("program", program)
      .eq("year", String(activeYear));

    if (data) setRecords(data);
  }, [program, activeYear]);

  useEffect(() => {
    loadIndirectFromDB();
  }, [program, activeYear]);

  useEffect(() => {
    if (program) loadData();
  }, [program, activeYear, loadData]);

  useEffect(() => {
    console.log("AI NARRATIVE:", generateAINarrative());
  }, [ploResults, indirectResults]);

  // FIX #3: Single source-of-truth — recalculate whenever plos OR records change
  useEffect(() => {
    if (plos.length > 0) {
      calculatePLO(records);
    }
  }, [plos, records]);

  useEffect(() => {
    console.log("REPORT DATA:", buildPILOsReportData());
  }, [ploResults, indirectResults]);

  const loadPreviousYearResults = useCallback(async () => {
    if (!program || !previousYearSelected) return;

    const { data } = await supabase
      .from("clo_records")
      .select("*")
      .eq("program", program)
      .eq("year", previousYearSelected);

    if (!data) return;

    const result = {};

    plos.forEach(p => {
      result[p.name] = 0;
    });

    data.forEach(i => {
      const [ploName, percent] = (i.mapping || "").split(":");
      const ach = Number(i.achievement) || 0;
      const w = Number(i.weight) || 1;
      const percentNum = Number(percent) || 0;

      const val = (ach * percentNum * w) / 10000;
      result[ploName] += val;
    });

    setPreviousYear(result);
  }, [program, previousYearSelected, plos]);

  useEffect(() => {
    loadPreviousYearResults();
  }, [loadPreviousYearResults]);

  /* ─── PLO CALCULATION ─── */
  const calculatePLO = (data) => {
    const result = {};

    // Initialise all known PLOs with zero
    plos.forEach(p => {
      result[p.name] = { total: 0, weight: 0 };
    });

    data.forEach(i => {
      const mappings = (i.mapping || "").split(",").filter(Boolean);
      mappings.forEach(m => {
        const [ploName, percent] = m.split(":");
        const percentNum = Number(percent) || 0;
        const ach = Number(i.achievement) || 0;
        const w = Number(i.weight) || 1;
        const val = (ach * percentNum * w) / 10000;

        if (!result[ploName]) {
          result[ploName] = { total: 0, weight: 0 };
        }
        result[ploName].total += val;
        result[ploName].weight += w;
      });
    });

    const final = {};
    Object.keys(result).forEach(k => {
      final[k] =
        result[k].weight === 0
          ? 0
          : Number((result[k].total / result[k].weight).toFixed(1));
    });

    setPloResults(final);
  };

  /* ─── PROGRAM FUNCTIONS ─── */
  const addProgram = async () => {
    if (!newProgram) return;
    await supabase.from("programs").insert([{ name: newProgram }]);
    setNewProgram("");
    loadPrograms();
  };

  const deleteProgram = async (id) => {
    await supabase.from("programs").delete().eq("id", id);
    loadPrograms();
  };

  const updateProgram = async (id) => {
    await supabase.from("programs")
      .update({ name: editedName })
      .eq("id", id);
    setEditingId(null);
    setEditedName("");
    loadPrograms();
  };

  /* ─── PLO FUNCTIONS ─── */
  const addPLO = async () => {
    if (!newPLO) return;
    await supabase.from("plos").insert([{ name: newPLO, program }]);
    setNewPLO("");
    loadPLOs();
  };

  const deletePLO = async (id) => {
    await supabase.from("plos").delete().eq("id", id);
    loadPLOs();
  };

  const updatePLO = async (id) => {
    await supabase.from("plos")
      .update({ name: editedPLOName })
      .eq("id", id);
    setEditingPLO(null);
    setEditedPLOName("");
    loadPLOs();
  };

  /* ─── CLO FUNCTIONS ─── */
  const addCLO = async () => {

    if (!selectedCourseId) {
      alert("Please select a Course");
      return;
    }

    const mappingString = `${plo}:100`;

    const { error } = await supabase.from("clo_records").insert([{
      description,
      achievement: Number(achievement),
      mapping: mappingString,
      weight: Number(weight) || 0,
      program,

      year: String(activeYear),
      course_id: Number(selectedCourseId)
    }]);

    if (error) {
      console.error("CLO INSERT ERROR:", error);
      alert("CLO was NOT added. Check console for error.");
      return;
    }

    setDescription("");
    setAchievement("");
    setPlo("");
    setWeight(0);
    loadData();

  };

  const deleteCLO = async (id) => {
    await supabase.from("clo_records").delete().eq("id", id);
    loadData();
  };

  /* ─── STATUS ─── */
  const getStatus = (v) => {
    if (v >= 70) return "🟢 Achieved";
    if (v >= 50) return "🟡 Acceptable";
    return "🔴 Needs Improvement";
  };

  const getImprovement = (cur, prev) =>
    ((cur - prev) >= 0 ? "+" : "") + (cur - prev).toFixed(1) + "%";

  /* ─── PDF ─── */
  // FIX #2: PDF now uses dynamic previousYear lookup with fallback to 0
  const exportPDF = () => {
    const doc = new jsPDF();
    doc.text("Academic Report", 20, 20);

    autoTable(doc, {
      head: [["PLO", "Current"]],
      body: Object.keys(ploResults).map(k => [
        k, ploResults[k] + "%"
      ])
    });

    autoTable(doc, {
      startY: doc.lastAutoTable.finalY + 10,
      head: [["PLO", "Prev", "Current", "Improvement"]],
      body: Object.keys(ploResults).map(k => {
        const prev = previousYear[k] ?? 0; // FIX #2: safe fallback for unknown PLOs
        return [
          k,
          prev + "%",
          ploResults[k] + "%",
          getImprovement(ploResults[k], prev)
        ];
      })
    });

    doc.save("report.pdf");
  };

  const chartData = Object.keys(ploResults).map(k => ({
    name: k,
    value: Number(ploResults[k])
  }));

  const cardStyle = {
    flex: 1,
    backgroundColor: "#fff",
    padding: "20px",
    borderRadius: "10px",
    boxShadow: "0 4px 10px rgba(0,0,0,0.2)",
    textAlign: "center"
  };

  const ploStatusCount = {
    achieved: 0,
    acceptable: 0,
    needsImprovement: 0
  };

  Object.values(ploResults).forEach(v => {
    if (v >= 70) ploStatusCount.achieved++;
    else if (v >= 50) ploStatusCount.acceptable++;
    else ploStatusCount.needsImprovement++;
  });

  const pieData = [
    { name: "Achieved", value: ploStatusCount.achieved },
    { name: "Acceptable", value: ploStatusCount.acceptable },
    { name: "Needs Improvement", value: ploStatusCount.needsImprovement }
  ];

  const comparisonData = Object.keys(ploResults).map(plo => ({
    name: plo,
    current: ploResults[plo] || 0,
    previous: previousYear[plo] || 0
  }));

  const handleExcelUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet);

      setExcelPreview(rows);
      setShowPreview(true);
    };

    reader.readAsBinaryString(file);
  };


  const handleIndirectExcelUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      console.log("ALL SHEETS:", workbook.SheetNames);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });
      // الصف الأول = العناوين
      const headers = rows[0];
      const dataRows = rows.slice(1);
      // تحويل الصفوف إلى Objects
      const parsedRows = dataRows.map(r => ({
        Program: r[headers.indexOf("Program")],
        PLO: r[headers.indexOf("PLO")],
        Indirect: r[headers.indexOf("Indirect")],
        Year: r[headers.indexOf("Year")]
      }));
      console.log("PARSED ROWS:", parsedRows);
      for (const row of parsedRows) {

        console.log("COMPARE:", {
          rowProgram: row.Program,
          program,
          rowYear: row.Year,
          activeYear
        });
        {

          const { error } = await supabase
            .from("indirect_assessment")
            .insert([{
              program,
              year: String(activeYear),
              plo: row.PLO,
              value: Number(row.Indirect)
            }]);

          if (error) {
            console.error("INSERT ERROR:", error);
          }
        }
      }
      alert("✅ Indirect Assessment saved successfully");
      loadIndirectFromDB();
    };

    reader.readAsBinaryString(file);
  };
  const [excelPreview, setExcelPreview] = useState([]);
  const [showPreview, setShowPreview] = useState(false);

  const confirmExcelImport = async () => {
    alert("✅ Confirm clicked (Import logic will go here)");
  };

  const downloadCLOTemplate = () => {
    const headers = [
      "Program",
      "CourseCode",
      "CLO_Description",
      "PLO",
      "Mapping",
      "Achievement",
      "Weight",
      "Year"
    ];

    // صف مثال إرشادي (اختياري)
    const exampleRow = [
      program || "PHD",
      "CS101",
      "Describe the CLO here",
      "plo1.1",
      100,
      75,
      1,
      activeYear || "2026"
    ];

    const worksheet = XLSX.utils.aoa_to_sheet([
      headers,
      exampleRow
    ]);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "CLO_Mapping");

    XLSX.writeFile(workbook, "CLO_Template.xlsx");
  };

  const loadIndirectFromDB = async () => {
    if (!program || !activeYear) return;

    const { data } = await supabase
      .from("indirect_assessment")
      .select("*")
      .eq("program", program)
      .eq("year", String(activeYear));

    const obj = {};
    data?.forEach(r => {
      obj[r.plo] = r.value;
    });

    setIndirectResults(obj);
  };

  const buildPILOsReportData = () => {
    const benchmark = 85;

    return Object.keys(ploResults).map((ploKey) => {
      const normalizedPLO = ploKey.replace(/\s+/g, "");
      const direct = ploResults[ploKey] || 0;
      const indirect = indirectResults[normalizedPLO] || 0;
      const weighted = Number((0.6 * direct + 0.4 * indirect).toFixed(1));

      let status = "Needs Improvement";
      if (weighted >= benchmark) status = "Achieved";
      else if (weighted >= 50) status = "In Progress";

      return {
        plo: ploKey,
        direct,
        indirect,
        weighted,
        benchmark,
        status
      };
    });
  };

  const renderPLOResultsTable = () => {
    const data = buildPILOsReportData();

    return (
      <table border="1" style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr>
            <th>PLO</th>
            <th>Direct (%)</th>
            <th>Indirect (%)</th>
            <th>Weighted (%)</th>
            <th>Benchmark</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          {data.map((row, index) => (
            <tr key={index}>
              <td>{row.plo}</td>
              <td>{row.direct}</td>
              <td>{row.indirect}</td>
              <td>{row.weighted}</td>
              <td>{row.benchmark}</td>
              <td>{row.status}</td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  const captureChartAsImage = async (chartId) => {
    const element = document.getElementById(chartId);
    if (!element) return null;

    const canvas = await html2canvas(element, {
      backgroundColor: "#ffffff"
    });

    return canvas.toDataURL("image/png");
  };

  const exportPILOsReportPDF = async () => {

    const doc = new jsPDF("p", "mm", "a4");

    // ---- Cover ----
    doc.setFontSize(16);
    doc.text("Program Intended Learning Outcomes (PILOs) Assessment Report", 15, 20);

    doc.setFontSize(11);
    doc.text(`University: ${reportMeta.university}`, 15, 30);
    doc.text(`College: ${reportMeta.college}`, 15, 36);
    doc.text(`Department: ${reportMeta.department}`, 15, 42);
    doc.text(`Program: ${reportMeta.program}`, 15, 48);
    doc.text(`Academic Year: ${reportMeta.academicYear}`, 15, 54);

    // ---- Table ----
    const tableData = buildPILOsReportData().map(row => ([
      row.plo,
      row.direct,
      row.indirect,
      row.weighted,
      row.benchmark,
      row.status
    ]));

    autoTable(doc, {
      startY: 65,
      head: [[
        "PLO",
        "Direct (%)",
        "Indirect (%)",
        "Weighted (%)",
        "Benchmark",
        "Status"
      ]],
      body: tableData
    });
    const aiText = generateAINarrative();

    let y = doc.lastAutoTable.finalY + 10;

    doc.setFontSize(12);
    doc.text("Introduction", 15, y);
    y += 6;
    doc.setFontSize(10);
    doc.text(aiText.introduction, 15, y, { maxWidth: 180 });

    y += 20;
    doc.setFontSize(12);
    doc.text("Results Interpretation", 15, y);
    y += 6;
    doc.setFontSize(10);
    doc.text(aiText.resultsInterpretation, 15, y, { maxWidth: 180 });

    y += 20;
    doc.setFontSize(12);
    doc.text("Conclusion", 15, y);
    y += 6;
    doc.setFontSize(10);
    doc.text(aiText.conclusion, 15, y, { maxWidth: 180 });

    y += 20;
    doc.setFontSize(12);
    doc.text("Recommendations", 15, y);
    y += 6;
    doc.setFontSize(10);
    doc.text(aiText.recommendations, 15, y, { maxWidth: 180 });

    // ===== Charts =====

    const barChartImg = await captureChartAsImage("plo-bar-chart");
    const pieChartImg = await captureChartAsImage("plo-pie-chart");

    if (barChartImg) {
      doc.addPage();
      doc.setFontSize(14);
      doc.text("PLO Achievement Bar Chart", 15, 20);
      doc.addImage(barChartImg, "PNG", 15, 30, 180, 90);
    }

    if (pieChartImg) {
      doc.addPage();
      doc.setFontSize(14);
      doc.text("PLO Status Distribution", 15, 20);
      doc.addImage(pieChartImg, "PNG", 40, 30, 120, 120);
    }

    // ---- Save ----
    doc.save(`PILOs_Report_${reportMeta.program}_${reportMeta.academicYear}.pdf`);
  };

  const buildAIReportSummary = () => {
    const data = buildPILOsReportData();

    const tableRows = [
      new TableRow({
        children: [
          "PLO",
          "Direct (%)",
          "Indirect (%)",
          "Weighted (%)",
          "Benchmark",
          "Status"
        ].map(text =>
          new TableCell({

            children: [new TextRun(text)]

          })
        )
      }),
      ...data.map(row =>
        new TableRow({
          children: [
            row.plo,
            row.direct.toString(),
            row.indirect.toString(),
            row.weighted.toString(),
            row.benchmark.toString(),
            row.status
          ].map(text =>
            new TableCell({
              children: [new Paragraph(text)]
            })
          )
        })
      )
    ];

    return {
      program: reportMeta.program,
      year: reportMeta.academicYear,
      totalPLOs: data.length,
      achieved: data.filter(d => d.status === "Achieved").length,
      inProgress: data.filter(d => d.status === "In Progress").length,
      needsImprovement: data.filter(d => d.status === "Needs Improvement").length
    };

  };

  const generateAIImprovementPlans = () => {
    const data = buildPILOsReportData();

    return data
      .filter(d => d.status !== "Achieved")
      .map(d => ({
        plo: d.plo,
        plan: `Improve ${d.plo} by reviewing course alignment, enhancing assessment methods, and introducing targeted learning activities addressing identified weaknesses.`
      }));
  };

  const base64ToUint8Array = (base64) => {
    const base64Data = base64.split(",")[1];
    const binaryString = atob(base64Data);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);

    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }

    return bytes;
  };

  const generateExecutiveSummary = () => {
    const data = Object.values(ploResults);

    const total = data.length;
    const achieved = data.filter(v => v >= 70).length;
    const acceptable = data.filter(v => v >= 50 && v < 70).length;
    const needsImprovement = data.filter(v => v < 50).length;
    const weakest = Object.entries(ploResults).sort((a, b) => a[1] - b[1])[0];
    const strongest = Object.entries(ploResults).sort((a, b) => b[1] - a[1])[0];

    const avg = total
      ? (data.reduce((a, b) => a + b, 0) / total).toFixed(1)
      : 0;

    return {
      text: `This report evaluates the achievement of Program Intended Learning Outcomes (PILOs) for the academic year ${reportMeta.academicYear}. 
Out of ${total} PLOs, ${achieved} were achieved, ${acceptable} were acceptable, and ${needsImprovement} require improvement. 
The overall average achievement is ${avg}%. 
The program is progressing adequately, however targeted improvement actions are required for low-performing outcomes.`
    };
  };

const generateAINarrative = () => {
  const data = buildPILOsReportData();
  


  const achieved = data.filter(d => d.status === "Achieved").length;
  const improvement = data.filter(d => d.status === "Needs Improvement").length;

  return {
    introduction: `This report presents an evaluation of the Program Intended Learning Outcomes (PILOs) for the academic year ${reportMeta.academicYear}. The assessment is based on both direct and indirect measures to ensure comprehensive analysis of student performance.`,

    resultsInterpretation: `The results indicate that ${achieved} PLOs were successfully achieved, while ${improvement} require further improvement. The variation in achievement levels reflects differences in course alignment and teaching effectiveness across the program.`,

    conclusion: `Overall, the program demonstrates acceptable performance; however, targeted improvements are required for lower-performing outcomes.`,

    recommendations: `It is recommended to enhance curriculum alignment, improve assessment strategies, and support students through focused interventions.`
  };
};
const generateDetailedRecommendations = () => {
  const data = buildPILOsReportData();

  return data.map(d => {
    let recommendation = "";

    if (d.weighted >= 85) {
      recommendation = "Maintain current teaching strategies.";
    } else if (d.weighted >= 70) {
      recommendation = "Enhance assessment methods.";
    } else if (d.weighted >= 50) {
      recommendation = "Improve course alignment.";
    } else {
      recommendation = "Urgent improvement required.";
    }

    return {
      plo: d.plo,
      text: recommendation,
      score: d.weighted
    };
  });
};
const exportPILOsReportWord = async () => {
  const data = buildPILOsReportData();
  const aiText = generateAINarrative();
  const summary = generateExecutiveSummary();
  const recs = generateDetailedRecommendations();
  
const safeTable = new Table({
  rows: [
    // Header
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph("PLO")] }),
        new TableCell({ children: [new Paragraph("Score (%)")] }),
        new TableCell({ children: [new Paragraph("Status")] }),
      ],
    }),

    // Data rows
    ...Object.entries(ploResults).map(([plo, value]) => {
      let status = "Needs Improvement";
      if (value >= 70) status = "Achieved";
      else if (value >= 50) status = "Acceptable";

      return new TableRow({
        children: [
          new TableCell({ children: [new Paragraph(plo)] }),
          new TableCell({ children: [new Paragraph(value + "%")] }),
          new TableCell({ children: [new Paragraph(status)] }),
        ],
      });
    }),
  ],
});

const ploTable = new Table({
  rows: [

    // ✅ Header
    new TableRow({
      children: ["PLO", "Direct (%)", "Indirect (%)", "Weighted (%)", "Status"]
        .map(header =>
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: header,
                    bold: true,
                    size: 24

                  })
                ]
              })
            ]
          })
        )
    }),

    // ✅ Data Rows
    ...Object.keys(ploResults).map(ploKey => {

      const direct = ploResults[ploKey] || 0;
      const indirect = indirectResults[ploKey.replace(/\s+/g, "")] || 0;
      const weighted = (0.6 * direct + 0.4 * indirect).toFixed(1);

      let status = "Needs Improvement";
      if (weighted >= 70) status = "Achieved";
      else if (weighted >= 50) status = "Acceptable";

      return new TableRow({
        children: [
          ploKey,
          `${direct}%`,
          `${indirect}%`,
          `${weighted}%`,
          status
        ].map(cell =>
          new TableCell({
            children: [new Paragraph(cell.toString())]
          })
        )
      });
    })
  ]
});

  const response = await fetch(logo);
  const logoBlob = await response.blob();
  const arrayBuffer = await logoBlob.arrayBuffer();
  const logoImage = new Uint8Array(arrayBuffer);

  const improvementPlans = generateAIImprovementPlans();

    const barChartImg = await captureChartAsImage("plo-bar-chart");
    const pieChartImg = await captureChartAsImage("plo-pie-chart");

    const children = [
      

new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [
    new TextRun({
      text: reportMeta.university,
      bold: true,
      size: 28
    })
  ]
}),

new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [
    new TextRun({
      text: reportMeta.college,
      size: 24
    })
  ]
}),

new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 300 },
  children: [
    new TextRun({
      text: reportMeta.department,
      size: 24
    })
  ]
}),

      
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 300 },
  children: [
    new TextRun({
      text: "EXECUTIVE SUMMARY",
      bold: true,
      size: 48
    })
  ]
}),
      new Paragraph(""),

    
new Paragraph({
  spacing: { after: 100 },
  children: [new TextRun({ text: `Total PLOs: ${Object.keys(ploResults).length}`, bold: true })]
}),

new Paragraph({
  children: [new TextRun(`Achieved: ${Object.values(ploResults).filter(v => v >= 70).length}`)]
}),

new Paragraph({
  children: [new TextRun(`Needs Improvement: ${Object.values(ploResults).filter(v => v < 50).length}`)]
}),

new Paragraph({
  children: [new TextRun(`Average Achievement: ${
    Object.values(ploResults).length
      ? (Object.values(ploResults).reduce((a,b)=>a+b,0) / Object.values(ploResults).length).toFixed(1)
      : 0
  }%`)]
}),

      // ✅ خط فاصل واضح
      new Paragraph("------------------------------"),

      // ✅ كسر صفحة أكيد
      new Paragraph({
        children: [],
        pageBreakBefore: true
      }),

      
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 400 },
  children: [
    new TextRun({
      text: "Program Intended Learning Outcomes (PILOs)",
      bold: true,
      size: 36
    }),
    new TextRun({
      text: "\nAssessment Report",
      size: 32
    })
  ]
}),


      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun(`University: ${reportMeta.university}`)]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun(`College: ${reportMeta.college}`)]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun(`Department: ${reportMeta.department}`)]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun(`Program: ${reportMeta.program}`)]
      }),
      new Paragraph({
        spacing: { after: 300 },
        children: [new TextRun(`Academic Year: ${reportMeta.academicYear}`)]
      }),

      new Paragraph(""),

      
new Paragraph({
  spacing: { before: 300, after: 150 },
  children: [
    new TextRun({
      text: "Introduction",
      bold: true,
      size: 28
    })
  ]
}),

      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun(aiText.introduction)]
      }),


new Paragraph({
  spacing: { before: 300, after: 150 },
  children: [
    new TextRun({
      text: "Results Interpretation",
      bold: true,
      size: 28
    })
  ]
}),

      new Paragraph(aiText.resultsInterpretation),
      
new Paragraph({
  children: [
    new TextRun({
      text: "PLO Recommendations",
      bold: true,
      size: 28
    })
  ]
}),

new Paragraph(""),



...recs.map(r => {
  let color = "000000";

  if (r.score >= 70) color = "008000";     // أخضر
  else if (r.score >= 50) color = "FFA500"; // برتقالي
  else color = "FF0000";                   // أحمر

  return new Paragraph({
    spacing: { after: 150 },
    children: [
      new TextRun("• "),
      new TextRun({
        text: `${r.plo} (${r.score}%)`,
        bold: true,
        color: color
      }),
      new TextRun(": "),
      new TextRun(r.text)
    ]
  });
}),




      
new Paragraph({
  spacing: { before: 300, after: 150 },
  children: [
    new TextRun({
      text: "conclusion",
      bold: true,
      size: 28
    })
  ]
}),

      new Paragraph(aiText.conclusion),

    

new Paragraph(""),

new Paragraph({
  children: [
    new TextRun({
      text: "PLO Assessment Matrix",
      bold: true,
      size: 30
    })
  ],
  spacing: { before: 300, after: 200 }
}),

safeTable,

      new Paragraph(""),

      // ===== Charts Section =====

      new Paragraph({
        children: [new TextRun({ text: "Assessment Charts", bold: true })]
      }),

    ];


    const doc = new Document({
      sections: [
        {
          children: children
        }
      ]
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `PILOs_Report_${reportMeta.program}_${reportMeta.academicYear}.docx`);
};

if (!user) {
  return (
    <div style={{ textAlign: "center", marginTop: "100px" }}>
      <h2>Login</h2>

      <input
        placeholder="Email"
        value={email}
        onChange={(e) => setEmail(e.target.value)}
      />

      <br /><br />

      <input
        type="password"
        placeholder="Password"
        value={password}
        onChange={(e) => setPassword(e.target.value)}
      />

      <br /><br />

      <button onClick={handleLogin}>
        Login
      </button>
    </div>
  );
}

if (!supabase) {
  return <div>Supabase not connected</div>;
}
return <h1>WORKING ✅</h1>;
    /* ─── UI ─── */
   }