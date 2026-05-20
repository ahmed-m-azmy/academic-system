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

  return <div>TEST ✅</div>;
}

  