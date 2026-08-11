"use client";

import { useState, useEffect } from "react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

// -------------------------------------------------------------
// 1. TYPES & CONFIGURATIONS
// -------------------------------------------------------------
interface MCQ {
  id: number;
  topic: string;
  question: string;
  options: string[];
  answer: string;
  explanation: string;
  tag?: string;
  customAudioUrl?: string;
}

const DESIGN_THEMES = {
  darkNavy: {
    name: "🌌 Gamma Dark Navy",
    bg: "0F172A",
    cardBg: "1E293B",
    titleColor: "38BDF8",
    textColor: "FFFFFF",
    subTextColor: "CBD5E1",
    correctBg: "065F46",
    correctColor: "34D399",
    border: "334155",
  },
  vintageCream: {
    name: "📜 Vintage Cream",
    bg: "FDF6E3",
    cardBg: "EEE8D5",
    titleColor: "B58900",
    textColor: "073642",
    subTextColor: "586E75",
    correctBg: "D3E4CD",
    correctColor: "2B580C",
    border: "D2C9B1",
  },
  minimalWhite: {
    name: "💎 Minimal Light",
    bg: "FFFFFF",
    cardBg: "F1F5F9",
    titleColor: "1D4ED8",
    textColor: "0F172A",
    subTextColor: "475569",
    correctBg: "DCFCE7",
    correctColor: "15803D",
    border: "E2E8F0",
  },
  emeraldNature: {
    name: "🌿 Emerald Forest",
    bg: "064E3B",
    cardBg: "047857",
    titleColor: "FDE047",
    textColor: "FFFFFF",
    subTextColor: "D1FAE5",
    correctBg: "022C22",
    correctColor: "34D399",
    border: "059669",
  },
  sunsetAmber: {
    name: "🔥 Sunset Amber",
    bg: "18181B",
    cardBg: "27272A",
    titleColor: "F97316",
    textColor: "FAFAFA",
    subTextColor: "A1A1AA",
    correctBg: "7C2D12",
    correctColor: "FDBA74",
    border: "3F3F46",
  },
};

const AVAILABLE_FONTS = [
  { id: "Nirmala UI", name: "✨ Nirmala UI (हिंदी एवं इंग्लिश - अनुशंसित)" },
  { id: "Segoe UI", name: "🔹 Segoe UI (मॉडर्न एवं साफ़)" },
  { id: "Trebuchet MS", name: "🔸 Trebuchet MS (बोल्ड एवं स्टाइल)" },
  { id: "Arial", name: "📄 Arial (क्लासिक सिंपल)" },
  { id: "Century Gothic", name: "💎 Century Gothic (प्रीमियम मिनिमल)" },
  { id: "Georgia", name: "📖 Georgia (एलिगेंट सेरिफ़)" },
];

const fontSizes = {
  small: { question: 15, option: 14, title: 12 },
  medium: { question: 18, option: 16, title: 13 },
  large: { question: 22, option: 18, title: 14 },
};

// -------------------------------------------------------------
// 2. HELPER & CLEANING FUNCTIONS
// -------------------------------------------------------------
const cleanText = (text: string) => {
  if (!text) return "";
  return text
    .replace(/[\*\_\`]/g, "")
    .replace(/^[🌍📖✅✔•\-\s]+/gu, "")
    .trim();
};

/**
 * 🛠️ DEVANAGARI MATRA REPAIR
 * Re-attaches isolated pre-base 'ि' (U+093F) matras without destroying word spaces.
 */
const fixMisplacedMatras = (text: string): string => {
  if (!text) return "";
  return text.replace(/\u093F([\u0904-\u0939\u0958-\u095F])/g, "$1\u093F");
};

/**
 * 📐 EXACT DOCUMENT LAYOUT FORMATTER
 * Formats PDF text into clean, structured lines for Questions, Options, and Answers.
 */
const formatDocumentStructure = (rawText: string): string => {
  if (!rawText) return "";

  let s = fixMisplacedMatras(rawText);

  // Clean up horizontal spacing line by line
  s = s
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n");

  // Newline before Questions (Q1, Q2, Q1 [, Question 1, प्रश्न 1)
  s = s.replace(/(?<!\n)\s*(Q\d{1,4}\b|Question\s*\d+|प्रश्न\s*\d+)/gi, "\n\n$1");

  // Separate Options (A), (B), (C), (D) onto individual lines if combined
  s = s.replace(/([^\n])\s*(\([A-Da-d]\))/g, "$1\n$2");

  // Separate Options formatted as A., B., C., D.
  s = s.replace(/([^\n])\s*\b([A-D])[\.\)]\s+(?=[^\n])/g, "$1\n($2) ");

  // Separate Answers and Status lines onto new lines
  s = s.replace(/([^\n])\s*(Your Answer:|Correct Answer:|Correct|Incorrect|Answer:|Ans:|उत्तर:|सही उत्तर:|व्याख्या:|Explanation:)/gi, "$1\n$2");

  // Remove stray orphan '(' characters on empty lines
  s = s.replace(/\n\s*\(\s*\n/g, "\n");    // Limit consecutive blank lines to at most 2   s = s.replace(/\n{3,}/g, "\n\n");    return s.trim(); };  const parseRawTextToMCQs = (rawText: string): MCQ[] => {   if (!rawText \vert{}\vert{} !rawText.trim()) return [];    const normalizedText = formatDocumentStructure(rawText);    let detectedTopic = "सामान्य ज्ञान (GK)";    const topicHeaderMatch = normalizedText.match(/^[^\n\d]*?([\u0900-\u097F\w\s]+?)(?:—\vert{}–\vert{}-\vert{}\d)/i);   if (topicHeaderMatch && topicHeaderMatch[1].trim()) {     const foundTopic = topicHeaderMatch[1].replace(/^[🌍📖✅✔•\-\s]+/gu, "").trim();     if (foundTopic.length > 2) detectedTopic = foundTopic;   }    const rawBlocks = normalizedText.split(/(?=\n\s*(?:Q\d{1,4}\vert{}प्रश्न\s*\d+\vert{}\b\d{1,3}\b)[\.\:\-\)\s\[]+)/gi);
  const parsedMcqs: MCQ[] = [];

  rawBlocks.forEach((block, idx) => {
    const trimmedBlock = block.trim();
    if (!trimmedBlock) return;

    const lines = trimmedBlock
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l.length > 0);

    if (lines.length === 0) return;

    let qText = "";
    let optA = "", optB = "", optC = "", optD = "";
    let ansLetter = "";
    let rawAnsLine = "";

    for (const line of lines) {
      const cleanL = line.replace(/^[🌍📖✅✔•\-\s]+/gu, "").trim();
      if (/^(?:Q\d{1,4}|प्रश्न\s*\d+|\b\d{1,3}\b)[\.\:\-\)\s\[]+/i.test(cleanL)) {
        qText = cleanL.replace(/^(?:Q\d{1,4}|प्रश्न\s*\d+|\b\d{1,3}\b)[\.\:\-\)\s\[]+\s*/i, "").trim();
        break;
      }
    }

    if (!qText && lines[0] && !/^\(?\b[A-D]\)?[\.\:\-\)\s]+/i.test(lines[0])) {
      qText = lines[0].replace(/^[🌍📖✅✔•\-\s]+/gu, "").trim();
    }

    lines.forEach((line) => {
      const cleanL = line.replace(/^[🌍📖✅✔•\-\s]+/gu, "").trim();

      if (/^\(?A\)?[\.\:\-\s]+/i.test(cleanL)) {
        optA = cleanL.replace(/^\(?A\)?[\.\:\-\s]*/i, "").trim();
      } else if (/^\(?B\)?[\.\:\-\s]+/i.test(cleanL)) {
        optB = cleanL.replace(/^\(?B\)?[\.\:\-\s]*/i, "").trim();
      } else if (/^\(?C\)?[\.\:\-\s]+/i.test(cleanL)) {
        optC = cleanL.replace(/^\(?C\)?[\.\:\-\s]*/i, "").trim();
      } else if (/^\(?D\)?[\.\:\-\s]+/i.test(cleanL)) {
        optD = cleanL.replace(/^\(?D\)?[\.\:\-\s]*/i, "").trim();
      } else if (
        cleanL.includes("Your Answer:") ||
        cleanL.includes("Correct Answer:") ||
        cleanL.includes("उत्तर") ||
        cleanL.toLowerCase().includes("answer")
      ) {
        rawAnsLine = cleanL;
        const letterMatch = cleanL.match(/(?:Your Answer:|Correct Answer:|उत्तर|Answer|Ans)[\:\-\s]*\(?([A-D])\)?/i);
        if (letterMatch) {
          ansLetter = letterMatch[1].toUpperCase();
        }
      }
    });

    const options = [optA.trim(), optB.trim(), optC.trim(), optD.trim()];

    let finalAns = options[0] || "";
    if (ansLetter === "A") finalAns = options[0];
    else if (ansLetter === "B") finalAns = options[1];
    else if (ansLetter === "C") finalAns = options[2];
    else if (ansLetter === "D") finalAns = options[3];
    else if (rawAnsLine) {
      const matched = options.find((o) => o && rawAnsLine.includes(o));
      if (matched) finalAns = matched;
    }

    let expText = "";
    const expIdx = trimmedBlock.search(/(?:व्याख्या|Explanation|Exp)[\:\-\s]*/i);
    if (expIdx !== -1) {
      const rawExp = trimmedBlock.slice(expIdx);
      expText = rawExp
        .replace(/^(?:व्याख्या|Explanation|Exp)[\:\-\s]*/i, "")
        .replace(/^[📖🌍✅✔•\-\s]+/gu, "")
        .trim();

      const ansInExp = expText.search(/(?:\n|\r|^)[📖🌍✅✔•\-\s]*(?:उत्तर|Answer|Ans)[\:\-\s]*/i);
      if (ansInExp !== -1) {
        expText = expText.slice(0, ansInExp).trim();
      }
    }

    if (qText && options.some((o) => o !== "")) {
      parsedMcqs.push({
        id: Date.now() + idx,
        topic: detectedTopic,
        question: qText,
        options: options,
        answer: finalAns,
        explanation: expText || "व्याख्या उपलब्ध नहीं है।",
        tag: "GK Set",
      });
    }
  });

  return parsedMcqs;
};

// -------------------------------------------------------------
// 3. MAIN REACT COMPONENT
// -------------------------------------------------------------
export default function Home() {
  const [activeTab, setActiveTab] = useState<"docTool" | "mcq">("docTool");

  // MCQ Generator States
  const [mcqList, setMcqList] = useState<MCQ[]>([
    {
      id: 1,
      topic: "सामान्य ज्ञान",
      question: "भारत का राष्ट्रीय जलीय जीव कौन-सा है?",
      options: ["मगरमच्छ", "गंगा नदी डॉल्फिन", "कछुआ", "व्हेल"],
      answer: "गंगा नदी डॉल्फिन",
      explanation:
        "गंगा नदी डॉल्फिन भारत का राष्ट्रीय जलीय जीव है। यह मुख्यतः गंगा-ब्रह्मपुत्र नदी तंत्र में पाई जाती है।",
      tag: "GK Practice Set",
    },
  ]);

  const [currentIndex, setCurrentIndex] = useState(0);
  const [pastedText, setPastedText] = useState("");
  const [showPasteBox, setShowPasteBox] = useState(false);
  const [brandingName, setBrandingName] = useState("My Study Library");
  const [selectedThemeKey, setSelectedThemeKey] =
    useState<keyof typeof DESIGN_THEMES>("vintageCream");

  const [selectedFont, setSelectedFont] = useState("Nirmala UI");
  const [fontSize, setFontSize] = useState<"small" | "medium" | "large">("medium");
  const [layoutMode, setLayoutMode] = useState<"grid" | "vertical">("grid");

  // Gemini API Voice States
  const [geminiApiKey, setGeminiApiKey] = useState("");
  const [isGeneratingAudio, setIsGeneratingAudio] = useState(false);
  const [isSpeaking, setIsSpeaking] = useState(false);

  // Document Exact Extractor States
  const [extractedDocText, setExtractedDocText] = useState("");
  const [extractedFileName, setExtractedFileName] = useState("");
  const [isExtractingDoc, setIsExtractingDoc] = useState(false);

  const currentMCQ = mcqList[currentIndex] || mcqList[0];
  const activeTheme = DESIGN_THEMES[selectedThemeKey];
  const currentFontSizes = fontSizes[fontSize];

  useEffect(() => {
    if (typeof window !== "undefined" && "speechSynthesis" in window) {
      window.speechSynthesis.cancel();
      setIsSpeaking(false);
    }
  }, [currentIndex]);

  const speakText = (text: string) => {
    if (typeof window !== "undefined" && "speechSynthesis" in window) {
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = "hi-IN";
      utterance.rate = 0.9;
      const voices = window.speechSynthesis.getVoices();
      const hindiVoice = voices.find((v) => v.lang.includes("hi"));
      if (hindiVoice) utterance.voice = hindiVoice;

      utterance.onstart = () => setIsSpeaking(true);
      utterance.onend = () => setIsSpeaking(false);
      utterance.onerror = () => setIsSpeaking(false);

      window.speechSynthesis.speak(utterance);
    }
  };

  const handleGenerateGeminiVoice = async () => {
    if (!currentMCQ?.question) {
      alert("कोई प्रश्न उपलब्ध नहीं है!");
      return;
    }

    setIsGeneratingAudio(true);

    const scriptText = `प्रश्न ${currentIndex + 1}. ${currentMCQ.question}. विकल्प A: ${currentMCQ.options[0] || ""}. विकल्प B: ${currentMCQ.options[1] || ""}. विकल्प C: ${currentMCQ.options[2] || ""}. विकल्प D: ${currentMCQ.options[3] || ""}. सही उत्तर है: ${currentMCQ.answer || ""}. व्याख्या: ${currentMCQ.explanation || ""}`;

    if (geminiApiKey.trim()) {
      try {
        const response = await fetch(
          `https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key=${geminiApiKey.trim()}`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              contents: [
                {
                  parts: [
                    {
                      text: `Summarize and format for speech narration: ${scriptText}`,
                    },
                  ],
                },
              ],
            }),
          }
        );

        if (!response.ok) {
          console.warn("Gemini API call returned non-200 status, using fallback TTS.");
        }
      } catch (err) {
        console.warn("Gemini API error, using fallback TTS.", err);
      }
    }

    speakText(scriptText);
    setIsGeneratingAudio(false);
  };

  const handleStopVoice = () => {
    if (typeof window !== "undefined" && "speechSynthesis" in window) {
      window.speechSynthesis.cancel();
      setIsSpeaking(false);
    }
  };

  const handleCustomAudioUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const audioUrl = URL.createObjectURL(file);
    updateCurrentMCQ("customAudioUrl", audioUrl);
    alert(`🎉 आपकी आवाज़ Q${currentIndex + 1} से जुड़ गई है!`);
  };

  const updateCurrentMCQ = (field: keyof MCQ, value: any) => {
    setMcqList((prevList) => {
      const updated = [...prevList];
      const cleanVal = typeof value === "string" ? cleanText(value) : value;
      updated[currentIndex] = { ...updated[currentIndex], [field]: cleanVal };
      return updated;
    });
  };

  const handleAddQuestion = () => {
    const newId = Date.now();
    const newMcq: MCQ = {
      id: newId,
      topic: cleanText(currentMCQ?.topic || "सामान्य ज्ञान"),
      question: "",
      options: ["", "", "", ""],
      answer: "",
      explanation: "",
      tag: "Practice Set",
    };

    setMcqList((prevList) => [...prevList, newMcq]);
    setCurrentIndex(mcqList.length);
  };

  const handleDeleteQuestion = (indexToDelete: number) => {
    if (mcqList.length === 1) {
      alert("कम से कम एक प्रश्न होना अनिवार्य है!");
      return;
    }
    setMcqList((prevList) =>
      prevList.filter((_, idx) => idx !== indexToDelete)
    );
    setCurrentIndex(Math.max(0, indexToDelete - 1));
  };

  const handleClearAll = () => {
    if (confirm("क्या आप सारा डेटा रीसेट करके नया सेट शुरू करना चाहते हैं?")) {
      setMcqList([
        {
          id: Date.now(),
          topic: "सामान्य ज्ञान",
          question: "",
          options: ["", "", "", ""],
          answer: "",
          explanation: "",
          tag: "Practice Set",
        },
      ]);
      setCurrentIndex(0);
    }
  };

  const handleCopyTextPaper = () => {
    let paperText = `📌 ${mcqList[0]?.topic || "महत्वपूर्ण प्रश्नोत्तरी"}\n\n`;

    mcqList.forEach((mcq, idx) => {
      paperText += `Q${idx + 1}. ${mcq.question}\n`;
      mcq.options.forEach((opt, oIdx) => {
        const letter = String.fromCharCode(65 + oIdx);
        paperText += `(${letter}) ${opt}\n`;
      });
      paperText += `✓ उत्तर: ${mcq.answer}\n`;
      if (mcq.explanation) paperText += `💡 व्याख्या: ${mcq.explanation}\n`;
      paperText += `-----------------------------------\n\n`;
    });

    navigator.clipboard.writeText(paperText);
    alert("🎉 पूरा प्रश्न-पत्र क्लिपबोर्ड में कॉपी हो गया है!");
  };

  const handleProcessPastedText = () => {
    if (!pastedText.trim()) {
      alert("कृपया बॉक्स में प्रश्न पेस्ट करें!");
      return;
    }

    const parsedMcqs = parseRawTextToMCQs(pastedText);

    if (parsedMcqs.length > 0) {
      setMcqList(parsedMcqs);
      setCurrentIndex(0);
      setPastedText("");
      setShowPasteBox(false);
      alert(`🎉 सफलता! व्याख्या के साथ पूरे ${parsedMcqs.length} प्रश्न लोड हो चुके हैं!`);
    } else {
      alert("पेस्ट किए गए टेक्स्ट को पढ़ने में समस्या आई। कृपया प्रश्न और उत्तर का फॉर्मेट जाँचें।");
    }
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data: any[] = XLSX.utils.sheet_to_json(ws);

      if (data.length === 0) {
        alert("फ़ाइल खाली है!");
        return;
      }

      const parsedMcqs: MCQ[] = data.map((row: any, idx: number) => ({
        id: Date.now() + idx,
        topic: cleanText(row.Topic || row.topic || "सामान्य ज्ञान"),
        question: cleanText(row.Question || row.question || ""),
        options: [
          cleanText(row.OptionA || row.optionA || row["Option A"] || ""),
          cleanText(row.OptionB || row.optionB || row["Option B"] || ""),
          cleanText(row.OptionC || row.optionC || row["Option C"] || ""),
          cleanText(row.OptionD || row.optionD || row["Option D"] || ""),
        ],
        answer: cleanText(row.Answer || row.answer || ""),
        explanation: cleanText(row.Explanation || row.explanation || ""),
        tag: "Practice Set",
      }));

      setMcqList(parsedMcqs);
      setCurrentIndex(0);
      alert(`🎉 सफलता! Excel से ${parsedMcqs.length} प्रश्न अपलोड हो चुके हैं।`);
    };

    reader.readAsBinaryString(file);
  };

  const handleDocxUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const arrayBuffer = evt.target?.result as ArrayBuffer;
      const result = await mammoth.extractRawText({ arrayBuffer });
      const rawText = result.value;

      if (!rawText.trim()) {
        alert("Word फ़ाइल खाली है!");
        return;
      }

      const parsedMcqs = parseRawTextToMCQs(rawText);

      if (parsedMcqs.length > 0) {
        setMcqList(parsedMcqs);
        setCurrentIndex(0);
        alert(`🎉 सफलता! Word (.docx) फ़ाइल से ${parsedMcqs.length} प्रश्न लोड हो गए हैं!`);
      } else {
        alert("Word फ़ाइल से प्रश्न पढ़ने में समस्या आई।");
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const handlePdfUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const pdfjsLib = await import("pdfjs-dist/build/pdf");
      pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

      const arrayBuffer = await file.arrayBuffer();
      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      const pdf = await loadingTask.promise;

      let fullText = "";

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();

        const items = (textContent.items as any[]).filter((it) => it.str !== undefined);

        const lineBuckets: { y: number; items: any[] }[] = [];
        for (const item of items) {
          const itemY = item.transform ? Math.round(item.transform[5]) : 0;
          let bucket = lineBuckets.find((b) => Math.abs(b.y - itemY) <= 4);
          if (!bucket) {
            bucket = { y: itemY, items: [] };
            lineBuckets.push(bucket);
          }
          bucket.items.push(item);
        }

        lineBuckets.sort((a, b) => b.y - a.y);

        let pageStr = "";
        for (const bucket of lineBuckets) {
          bucket.items.sort((a, b) => (a.transform ? a.transform[4] : 0) - (b.transform ? b.transform[4] : 0));

          let lineText = "";
          let prevXEnd = null;

          for (const item of bucket.items) {
            const x = item.transform ? item.transform[4] : 0;
            const width = item.width || (item.str ? item.str.length * 5 : 0);

            if (prevXEnd !== null) {
              const gap = x - prevXEnd;
              if (gap > 3 && !lineText.endsWith(" ") && !item.str.startsWith(" ")) {
                lineText += " ";
              }
            }
            lineText += item.str;
            prevXEnd = x + width;
          }

          if (lineText.trim()) {
            pageStr += lineText.trim() + "\n";
          }
        }

        fullText += pageStr + "\n";
      }

      if (!fullText.trim()) {
        alert("यह एक इमेज/स्कैन्ड PDF है।");
        return;
      }

      const parsedMcqs = parseRawTextToMCQs(fullText);

      if (parsedMcqs.length > 0) {
        setMcqList(parsedMcqs);
        setCurrentIndex(0);
        alert(`🎉 सफलता! PDF फ़ाइल से सभी ${parsedMcqs.length} प्रश्न लोड हो चुके हैं!`);
      } else {
        alert("PDF फ़ाइल से प्रश्न पढ़ने में समस्या आई।");
      }
    } catch (err) {
      console.error("PDF Parsing Error:", err);
      alert("PDF फ़ाइल प्रोसेस करने में समस्या आई!");
    }
  };

  // -------------------------------------------------------------
  // 📄 ACCURATE SPATIAL DOCUMENT EXTRACTOR
  // -------------------------------------------------------------
  const handleExtractExactDocumentText = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsExtractingDoc(true);
    setExtractedFileName(file.name.replace(/\.[^/.]+$/, ""));

    try {
      let rawExtractedText = "";
      const extension = file.name.split(".").pop()?.toLowerCase();

      if (extension === "pdf") {
        const pdfjsLib = await import("pdfjs-dist/build/pdf");
        pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

        let fullDocText = "";

        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();

          const items = (textContent.items as any[]).filter((it) => it.str !== undefined);

          // Group items by vertical Y coordinate
          const lineBuckets: { y: number; items: any[] }[] = [];
          for (const item of items) {
            const itemY = item.transform ? Math.round(item.transform[5]) : 0;
            let bucket = lineBuckets.find((b) => Math.abs(b.y - itemY) <= 4);
            if (!bucket) {
              bucket = { y: itemY, items: [] };
              lineBuckets.push(bucket);
            }
            bucket.items.push(item);
          }

          // Sort lines Top -> Bottom (Y decreases downwards)
          lineBuckets.sort((a, b) => b.y - a.y);

          let pageLines = [];
          for (const bucket of lineBuckets) {
            // Sort items in line Left -> Right (X coordinate)
            bucket.items.sort((a, b) => (a.transform ? a.transform[4] : 0) - (b.transform ? b.transform[4] : 0));

            let lineStr = "";
            let prevXEnd = null;

            for (const item of bucket.items) {
              const x = item.transform ? item.transform[4] : 0;
              const width = item.width || (item.str ? item.str.length * 5 : 0);

              if (prevXEnd !== null) {
                const gap = x - prevXEnd;
                // Add space if there is a gap between adjacent text chunks on the line
                if (gap > 3 && !lineStr.endsWith(" ") && !item.str.startsWith(" ")) {
                  lineStr += " ";
                }
              }

              lineStr += item.str;
              prevXEnd = x + width;
            }

            if (lineStr.trim()) {
              pageLines.push(lineStr.trim());
            }
          }

          fullDocText += `=== पृष्ठ ${i} ===\n` + pageLines.join("\n") + "\n\n";
        }

        rawExtractedText = fullDocText;
      } else if (extension === "docx") {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        rawExtractedText = result.value;
      } else if (extension === "txt") {
        rawExtractedText = await file.text();
      } else {
        alert("केवल PDF, Word (.docx) या Text (.txt) फ़ाइल अपलोड करें!");
        setIsExtractingDoc(false);
        return;
      }

      if (rawExtractedText.trim()) {
        const structuredText = formatDocumentStructure(rawExtractedText);
        setExtractedDocText(structuredText);
        alert(`🎉 सफलता! '${file.name}' से पूरी तरह व्यवस्थित टेक्स्ट निकाल लिया गया है!`);
      } else {
        alert("फ़ाइल में से टेक्स्ट नहीं पढ़ा जा सका। फ़ाइल खाली या स्कैन्ड इमेज हो सकती है।");
      }
    } catch (err) {
      console.error("Document Extraction Error:", err);
      alert("दस्तावेज़ से टेक्स्ट निकालने में समस्या आई!");
    } finally {
      setIsExtractingDoc(false);
    }
  };

  const handleDownloadExtractedTxt = () => {
    if (!extractedDocText.trim()) return;
    const blob = new Blob([extractedDocText], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${extractedFileName || "Extracted_Text"}_Copy.txt`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadExtractedDoc = () => {
    if (!extractedDocText.trim()) return;
    const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Document Copy</title></head><body>";
    const footer = "</body></html>";
    const html = header + "<div style='font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; white-space: pre-wrap;'>" + extractedDocText.replace(/\n/g, "<br>") + "</div>" + footer;

    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${extractedFileName || "Document"}_Exact_Copy.doc`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleSendExtractedToMcq = () => {
    if (!extractedDocText.trim()) return;
    setPastedText(extractedDocText);
    setActiveTab("mcq");
    setShowPasteBox(true);
    alert("🎉 निकाला गया टेक्स्ट MCQ पेस्ट बॉक्स में भेज दिया गया है! अब 'Load Questions' बटन दबाएँ।");
  };

  // -------------------------------------------------------------
  // PPT EXPORT ENGINE
  // -------------------------------------------------------------
  const handleDownloadPPT = async () => {
    const validMcqs = mcqList.filter((m) => m.question.trim() !== "");

    if (validMcqs.length === 0) {
      alert("कृपया कम से कम एक प्रश्न दर्ज करें!");
      return;
    }

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";
    const theme = DESIGN_THEMES[selectedThemeKey];

    const PPT_FONT = selectedFont;

    const mainTopic = validMcqs[0]?.topic || "महत्वपूर्ण प्रश्नोत्तरी";

    const addBranding = (slide: pptxgen.Slide) => {
      if (brandingName.trim()) {
        slide.addText(`📢 ${brandingName}`, {
          x: 0.8,
          y: 5.2,
          w: 8.4,
          h: 0.3,
          fontSize: 10,
          bold: true,
          fontFace: PPT_FONT,
          color: theme.subTextColor,
          align: "right",
        });
      }
    };

    // COVER SLIDE
    const coverSlide = pptx.addSlide();
    coverSlide.background = { color: theme.bg };

    coverSlide.addShape("roundRect" as any, {
      x: 0.8,
      y: 0.6,
      w: 8.4,
      h: 4.4,
      fill: { color: theme.cardBg },
      line: { color: theme.titleColor, width: 2 },
    });

    coverSlide.addText(mainTopic, {
      x: 1.2,
      y: 1.2,
      w: 7.6,
      h: 1.0,
      fontSize: 32,
      bold: true,
      fontFace: PPT_FONT,
      color: theme.titleColor,
      align: "left",
    });

    coverSlide.addText("बहुविकल्पीय प्रश्नोत्तरी एवं विस्तृत व्याख्या सेट", {
      x: 1.2,
      y: 2.3,
      w: 7.6,
      h: 0.5,
      fontSize: 18,
      fontFace: PPT_FONT,
      color: theme.subTextColor,
      align: "left",
    });

    coverSlide.addText(`कुल प्रश्न: ${validMcqs.length} • Gamma Presentation`, {
      x: 1.2,
      y: 3.6,
      w: 7.6,
      h: 0.5,
      fontSize: 15,
      bold: true,
      fontFace: PPT_FONT,
      color: theme.correctColor,
    });
    addBranding(coverSlide);

    // MCQ SLIDES
    for (let i = 0; i < validMcqs.length; i++) {
      const mcq = validMcqs[i];

      const slide1 = pptx.addSlide();
      slide1.background = { color: theme.bg };

      slide1.addShape("roundRect" as any, {
        x: 0.8,
        y: 0.4,
        w: 8.4,
        h: 0.5,
        fill: { color: theme.cardBg },
      });

      slide1.addText(`विषय: ${mcq.topic} ${mcq.tag ? `[${mcq.tag}]` : ""}`, {
        x: 1.0,
        y: 0.4,
        w: 5.5,
        h: 0.5,
        fontSize: currentFontSizes.title,
        fontFace: PPT_FONT,
        color: theme.titleColor,
        bold: true,
      });

      slide1.addText(`प्रश्न ${i + 1} / ${validMcqs.length}`, {
        x: 6.6,
        y: 0.4,
        w: 2.4,
        h: 0.5,
        fontSize: currentFontSizes.title,
        fontFace: PPT_FONT,
        color: theme.subTextColor,
        bold: true,
        align: "right",
      });

      slide1.addShape("roundRect" as any, {
        x: 0.8,
        y: 1.0,
        w: 8.4,
        h: 1.3,
        fill: { color: theme.cardBg },
        line: { color: theme.titleColor, width: 1.5 },
      });

      slide1.addText(`Q${i + 1}.  ${mcq.question}`, {
        x: 1.0,
        y: 1.0,
        w: 8.0,
        h: 1.3,
        fontSize: currentFontSizes.question,
        bold: true,
        fontFace: PPT_FONT,
        color: theme.textColor,
        valign: "middle",
      });

      if (layoutMode === "grid") {
        const optPositions = [
          { x: 0.8, y: 2.5 },
          { x: 5.1, y: 2.5 },
          { x: 0.8, y: 3.8 },
          { x: 5.1, y: 3.8 },
        ];

        mcq.options.forEach((opt, optIdx) => {
          const pos = optPositions[optIdx];
          const optionLetter = String.fromCharCode(65 + optIdx);

          slide1.addShape("roundRect" as any, {
            x: pos.x,
            y: pos.y,
            w: 4.1,
            h: 1.1,
            fill: { color: theme.cardBg },
            line: { color: theme.border, width: 1 },
          });

          slide1.addText(`(${optionLetter})  ${opt}`, {
            x: pos.x + 0.2,
            y: pos.y,
            w: 3.7,
            h: 1.1,
            fontSize: currentFontSizes.option,
            bold: true,
            fontFace: PPT_FONT,
            color: theme.textColor,
            valign: "middle",
          });
        });
      } else {
        const optY = [2.4, 3.1, 3.8, 4.5];
        mcq.options.forEach((opt, optIdx) => {
          const y = optY[optIdx];
          const optionLetter = String.fromCharCode(65 + optIdx);

          slide1.addShape("roundRect" as any, {
            x: 0.8,
            y: y,
            w: 8.4,
            h: 0.6,
            fill: { color: theme.cardBg },
            line: { color: theme.border, width: 1 },
          });

          slide1.addText(`(${optionLetter})  ${opt}`, {
            x: 1.1,
            y: y,
            w: 7.8,
            h: 0.6,
            fontSize: currentFontSizes.option,
            bold: true,
            fontFace: PPT_FONT,
            color: theme.textColor,
            valign: "middle",
          });
        });
      }

      addBranding(slide1);

      // ANSWER SLIDE
      const slide2 = pptx.addSlide();
      slide2.background = { color: theme.bg };

      slide2.addShape("roundRect" as any, {
        x: 0.8,
        y: 0.4,
        w: 8.4,
        h: 0.5,
        fill: { color: theme.cardBg },
      });

      slide2.addText(`विषय: ${mcq.topic} • प्रश्न ${i + 1} (उत्तर एवं व्याख्या)`, {
        x: 1.0,
        y: 0.4,
        w: 8.0,
        h: 0.5,
        fontSize: currentFontSizes.title,
        fontFace: PPT_FONT,
        color: theme.titleColor,
        bold: true,
      });

      slide2.addShape("roundRect" as any, {
        x: 0.8,
        y: 1.0,
        w: 8.4,
        h: 1.0,
        fill: { color: theme.correctBg },
        line: { color: theme.correctColor, width: 2 },
      });

      slide2.addText(`✓ सही उत्तर: ${mcq.answer}`, {
        x: 1.1,
        y: 1.0,
        w: 7.8,
        h: 1.0,
        fontSize: currentFontSizes.question,
        bold: true,
        fontFace: PPT_FONT,
        color: theme.correctColor,
        valign: "middle",
      });

      slide2.addShape("roundRect" as any, {
        x: 0.8,
        y: 2.2,
        w: 8.4,
        h: 2.9,
        fill: { color: theme.cardBg },
        line: { color: theme.border, width: 1 },
      });

      slide2.addText("विस्तृत व्याख्या:", {
        x: 1.1,
        y: 2.4,
        w: 7.8,
        h: 0.4,
        fontSize: 16,
        bold: true,
        fontFace: PPT_FONT,
        color: theme.titleColor,
      });

      slide2.addText(mcq.explanation, {
        x: 1.1,
        y: 2.9,
        w: 7.8,
        h: 2.0,
        fontSize: currentFontSizes.option,
        fontFace: PPT_FONT,
        color: theme.textColor,
        valign: "top",
      });
      addBranding(slide2);
    }

    // SUMMARY SLIDE
    const summarySlide = pptx.addSlide();
    summarySlide.background = { color: theme.bg };

    summarySlide.addText("📌 त्वरित उत्तर कुंजी (Answer Key Summary)", {
      x: 0.8,
      y: 0.4,
      w: 8.4,
      h: 0.5,
      fontSize: 20,
      bold: true,
      fontFace: PPT_FONT,
      color: theme.titleColor,
    });

    const rows: any[] = [
      [
        { text: "क्र.", options: { bold: true, fontFace: PPT_FONT, fill: theme.cardBg, color: theme.titleColor } },
        { text: "प्रश्न (Question)", options: { bold: true, fontFace: PPT_FONT, fill: theme.cardBg, color: theme.titleColor } },
        { text: "सही उत्तर", options: { bold: true, fontFace: PPT_FONT, fill: theme.cardBg, color: theme.correctColor } },
      ],
    ];

    validMcqs.forEach((mcq, idx) => {
      rows.push([
        { text: `Q${idx + 1}`, options: { bold: true, fontFace: PPT_FONT, color: theme.textColor } },
        { text: mcq.question.slice(0, 45) + "...", options: { fontFace: PPT_FONT, color: theme.subTextColor } },
        { text: mcq.answer, options: { bold: true, fontFace: PPT_FONT, color: theme.correctColor } },
      ]);
    });

    summarySlide.addTable(rows, {
      x: 0.8,
      y: 1.0,
      w: 8.4,
      colW: [0.8, 5.2, 2.4],
      border: { pt: 1, color: theme.border },
      fill: theme.bg,
      fontSize: 13,
      autoPage: true,
    });
    addBranding(summarySlide);

    await pptx.writeFile({
      fileName: `${mainTopic}_Presentation.pptx`,
    });
  };

  return (
    <main className="min-h-screen bg-slate-950 text-white">
      {/* Header */}
      <header className="border-b border-white/10 bg-slate-900 sticky top-0 z-50 shadow-md">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
          <div>
            <h1 className="text-xl font-bold md:text-2xl text-transparent bg-clip-text bg-gradient-to-r from-blue-400 via-purple-400 to-yellow-400">
              MCQ Studio & Document Converter
            </h1>
            <p className="text-xs text-slate-400">
              Exact Document Copy Generator • MCQ PPT Builder • Gemini Audio
            </p>
          </div>

          {/* TAB SWITCHER */}
          <div className="flex bg-slate-800 p-1 rounded-xl border border-white/10">
            <button
              onClick={() => setActiveTab("docTool")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition ${
                activeTab === "docTool"
                  ? "bg-purple-600 text-white shadow-lg"
                  : "text-slate-400 hover:text-white"
              }`}
            >
              📄 Exact Document Copy Tool
            </button>
            <button
              onClick={() => setActiveTab("mcq")}
              className={`px-4 py-2 text-xs font-bold rounded-lg transition ${
                activeTab === "mcq"
                  ? "bg-blue-600 text-white shadow-lg"
                  : "text-slate-400 hover:text-white"
              }`}
            >
              🎯 MCQ PPT Studio
            </button>
          </div>

          <div className="hidden md:flex items-center gap-3">
            <button
              onClick={handleCopyTextPaper}
              className="rounded-lg bg-blue-600/30 px-3 py-2 text-xs font-bold text-blue-400 border border-blue-500/40 hover:bg-blue-600/50"
            >
              📋 Copy Paper
            </button>
            <button
              onClick={handleClearAll}
              className="rounded-lg bg-red-600/20 px-3 py-2 text-xs font-bold text-red-400 border border-red-500/40 hover:bg-red-600/40"
            >
              🧹 Clear
            </button>
          </div>
        </div>
      </header>

      {/* DYNAMIC CONTENT */}
      {activeTab === "docTool" ? (
        /* EXACT DOCUMENT COPY GENERATOR TAB */
        <div className="mx-auto max-w-6xl px-6 py-8">
          <div className="rounded-2xl border border-purple-500/30 bg-slate-900 p-6 shadow-2xl">
            <div className="mb-6 flex flex-col md:flex-row md:items-center justify-between border-b border-white/10 pb-4 gap-4">
              <div>
                <h2 className="text-xl font-bold text-purple-400 flex items-center gap-2">
                  📄 Exact Document Copy Generator (हूबहू कॉपी टूल)
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  हिंदी या अंग्रेज़ी की PDF, Word (.docx) या Text (.txt) फ़ाइल अपलोड करें और उसका पूरा टेक्स्ट प्राप्त करें।
                </p>
              </div>

              {extractedDocText && (
                <div className="flex flex-wrap gap-2">
                  <button
                    onClick={() => {
                      navigator.clipboard.writeText(extractedDocText);
                      alert("🎉 पूरा टेक्स्ट क्लिपबोर्ड में कॉपी हो गया!");
                    }}
                    className="rounded-lg bg-blue-600 px-3 py-2 text-xs font-bold hover:bg-blue-500 shadow"
                  >
                    📋 Copy Text
                  </button>
                  <button
                    onClick={handleDownloadExtractedTxt}
                    className="rounded-lg bg-green-600 px-3 py-2 text-xs font-bold hover:bg-green-500 shadow"
                  >
                    💾 Save as .txt
                  </button>
                  <button
                    onClick={handleDownloadExtractedDoc}
                    className="rounded-lg bg-yellow-600 px-3 py-2 text-xs font-bold text-slate-950 hover:bg-yellow-500 shadow"
                  >
                    📝 Save as Word (.doc)
                  </button>
                  <button
                    onClick={handleSendExtractedToMcq}
                    className="rounded-lg bg-purple-600 px-3 py-2 text-xs font-bold hover:bg-purple-500 shadow"
                  >
                    ⚡ Convert to MCQ PPT
                  </button>
                </div>
              )}
            </div>

            <div className="mb-6 rounded-2xl border-2 border-dashed border-purple-500/40 bg-purple-950/10 p-8 text-center">
              <div className="mx-auto mb-3 flex h-16 w-16 items-center justify-center rounded-full bg-purple-600/20 text-3xl text-purple-400">
                📁
              </div>
              <h3 className="text-base font-bold text-slate-200 mb-1">
                यहाँ हिंदी या अंग्रेज़ी फ़ाइल अपलोड करें
              </h3>
              <p className="text-xs text-slate-400 mb-4">
                समर्थित फ़ाइल फॉर्मेट: PDF (.pdf), Word (.docx), Plain Text (.txt)
              </p>
              <label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-purple-600 px-6 py-3 text-sm font-bold text-white shadow-lg transition hover:bg-purple-500">
                {isExtractingDoc ? "⏳ टेक्स्ट निकाला जा रहा है..." : "📂 Select PDF / Word File"}
                <input
                  type="file"
                  accept=".pdf, .docx, .txt"
                  onChange={handleExtractExactDocumentText}
                  disabled={isExtractingDoc}
                  className="hidden"
                />
              </label>
            </div>

            {extractedDocText ? (
              <div>
                <div className="mb-2 flex items-center justify-between text-xs font-bold text-slate-300">
                  <span>
                    📄 निकाला गया टेक्स्ट ({extractedFileName || "Document"}):
                  </span>
                  <span className="text-purple-400">
                    कुल शब्द: {extractedDocText.split(/\s+/).filter(Boolean).length} | कुल अक्षर: {extractedDocText.length}
                  </span>
                </div>
                <textarea
                  value={extractedDocText}
                  onChange={(e) => setExtractedDocText(e.target.value)}
                  className="h-[500px] w-full rounded-2xl border border-white/10 bg-slate-950 p-4 font-mono text-sm leading-relaxed text-slate-200 outline-none focus:border-purple-500 shadow-inner"
                  placeholder="यहाँ निकाला हुआ टेक्स्ट दिखेगा..."
                />
              </div>
            ) : (
              <div className="rounded-xl border border-white/5 bg-slate-950/50 p-12 text-center text-xs text-slate-500">
                💡 फ़ाइल अपलोड करते ही यहाँ पूरा टेक्स्ट आ जाएगा। आप उस टेक्स्ट में संपादन (Edit) भी कर सकते हैं।
              </div>
            )}
          </div>
        </div>
      ) : (
        /* MCQ PPT STUDIO TAB */
        <div className="mx-auto grid max-w-7xl gap-6 px-6 py-8 lg:grid-cols-2">
          {/* LEFT PANEL */}
          <section className="rounded-2xl border border-white/10 bg-slate-900 p-6 shadow-2xl">
            {/* GEMINI API KEY INPUT */}
            <div className="mb-5 rounded-xl border border-blue-500/30 bg-blue-950/20 p-3.5">
              <label className="block text-xs font-bold text-blue-300 mb-1">
                🔑 Gemini API Key (Paste Your Copied Key Here)
              </label>
              <input
                type="password"
                value={geminiApiKey}
                onChange={(e) => setGeminiApiKey(e.target.value)}
                className="w-full rounded-lg border border-white/10 bg-slate-900 px-3 py-2 text-xs text-slate-200 outline-none focus:border-blue-500"
                placeholder="Paste AI Studio API Key (AIzaSy...)"
              />
            </div>

            {/* BRANDING & WATERMARK INPUT */}
            <div className="mb-5 rounded-xl border border-white/10 bg-slate-800/80 p-3.5">
              <label className="block text-xs font-bold text-yellow-400 mb-1">
                🏷️ Channel / Branding Watermark (PPT Footer)
              </label>
              <input
                value={brandingName}
                onChange={(e) => setBrandingName(e.target.value)}
                className="w-full rounded-lg border border-white/10 bg-slate-900 px-3 py-2 text-xs text-slate-200 outline-none focus:border-yellow-500"
                placeholder="अपनी कोचिंग/यूट्यूब का नाम लिखें..."
              />
            </div>

            {/* THEME SELECTOR BOX */}
            <div className="mb-5 rounded-xl border border-white/10 bg-slate-800/80 p-4">
              <h3 className="mb-3 text-sm font-bold text-blue-400">🎨 Choose Presentation Design Theme</h3>
              <div className="grid grid-cols-2 gap-2.5 sm:grid-cols-3">
                {(Object.keys(DESIGN_THEMES) as Array<keyof typeof DESIGN_THEMES>).map((key) => (
                  <button
                    key={key}
                    onClick={() => setSelectedThemeKey(key)}
                    className={`rounded-xl border p-2.5 text-xs font-bold transition text-left ${
                      selectedThemeKey === key
                        ? "border-blue-500 bg-blue-600/20 text-blue-400 shadow-md ring-2 ring-blue-500/30"
                        : "border-white/10 bg-slate-800 text-slate-400 hover:bg-slate-700"
                    }`}
                  >
                    {DESIGN_THEMES[key].name}
                  </button>
                ))}
              </div>
            </div>

            {/* FONT SELECTOR */}
            <div className="mb-5 rounded-xl border border-white/10 bg-slate-800/80 p-4">
              <label className="block text-xs font-bold text-green-400 mb-2">
                🔤 Choose PPT Slide Font Style (फ़ॉन्ट शैली चुनें)
              </label>
              <select
                value={selectedFont}
                onChange={(e) => setSelectedFont(e.target.value)}
                className="w-full rounded-xl border border-white/10 bg-slate-900 px-3 py-2.5 text-xs font-bold text-slate-200 outline-none focus:border-green-500 shadow-sm"
              >
                {AVAILABLE_FONTS.map((font) => (
                  <option key={font.id} value={font.id}>
                    {font.name}
                  </option>
                ))}
              </select>
            </div>

            {/* FONT SIZE & LAYOUT CONTROLS */}
            <div className="mb-5 grid grid-cols-2 gap-3 rounded-xl border border-white/10 bg-slate-800/60 p-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-300 mb-1.5">📏 Text Font Size</label>
                <div className="flex gap-1.5">
                  {(["small", "medium", "large"] as const).map((sz) => (
                    <button
                      key={sz}
                      onClick={() => setFontSize(sz)}
                      className={`flex-1 rounded-lg py-1.5 text-xs font-bold capitalize transition ${
                        fontSize === sz
                          ? "bg-blue-600 text-white"
                          : "bg-slate-900 text-slate-400 hover:bg-slate-700"
                      }`}
                    >
                      {sz}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-300 mb-1.5">🎛️ Options Layout</label>
                <div className="flex gap-1.5">
                  <button
                    onClick={() => setLayoutMode("grid")}
                    className={`flex-1 rounded-lg py-1.5 text-xs font-bold transition ${
                      layoutMode === "grid"
                        ? "bg-blue-600 text-white"
                        : "bg-slate-900 text-slate-400 hover:bg-slate-700"
                    }`}
                  >
                    2x2 Grid
                  </button>
                  <button
                    onClick={() => setLayoutMode("vertical")}
                    className={`flex-1 rounded-lg py-1.5 text-xs font-bold transition ${
                      layoutMode === "vertical"
                        ? "bg-blue-600 text-white"
                        : "bg-slate-900 text-slate-400 hover:bg-slate-700"
                    }`}
                  >
                    Vertical
                  </button>
                </div>
              </div>
            </div>

            {/* BULK UPLOAD OPTIONS */}
            <div className="mb-4">
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-xs font-bold text-slate-300">📥 Choose Input Method</h3>
                <button
                  onClick={() => setShowPasteBox(!showPasteBox)}
                  className="text-xs font-bold text-yellow-400 hover:underline flex items-center gap-1"
                >
                  📋 {showPasteBox ? "Hide Paste Box" : "📋 Direct Copy-Paste Text Area"}
                </button>
              </div>

              {showPasteBox ? (
                <div className="mb-4 rounded-xl border border-yellow-500/40 bg-yellow-950/20 p-4">
                  <h4 className="font-bold text-xs text-yellow-400 mb-1">
                    📋 Paste Your Questions Below:
                  </h4>
                  <textarea
                    value={pastedText}
                    onChange={(e) => setPastedText(e.target.value)}
                    className="h-32 w-full rounded-lg border border-white/10 bg-slate-800 p-3 text-xs text-slate-200 outline-none focus:border-yellow-500"
                    placeholder="यहाँ अपने प्रश्न पेस्ट करें..."
                  />
                  <button
                    onClick={handleProcessPastedText}
                    className="mt-2 w-full rounded-lg bg-yellow-600 py-2.5 text-xs font-bold text-slate-950 transition hover:bg-yellow-500"
                  >
                    ✨ Load Questions from Pasted Text
                  </button>
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                  <div className="rounded-xl border border-dashed border-blue-500/40 bg-blue-950/20 p-2.5">
                    <h3 className="font-bold text-[10px] text-blue-400 mb-1">📁 Excel / CSV</h3>
                    <input
                      type="file"
                      accept=".xlsx, .xls, .csv"
                      onChange={handleExcelUpload}
                      className="w-full cursor-pointer rounded bg-slate-800 p-1 text-[9px] text-slate-300 file:mr-1 file:rounded file:border-0 file:bg-blue-600 file:px-1.5 file:py-0.5 file:text-[9px] file:font-semibold file:text-white"
                    />
                  </div>

                  <div className="rounded-xl border border-dashed border-purple-500/40 bg-purple-950/20 p-2.5">
                    <h3 className="font-bold text-[10px] text-purple-400 mb-1">📄 Word (.docx)</h3>
                    <input
                      type="file"
                      accept=".docx"
                      onChange={handleDocxUpload}
                      className="w-full cursor-pointer rounded bg-slate-800 p-1 text-[9px] text-slate-300 file:mr-1 file:rounded file:border-0 file:bg-purple-600 file:px-1.5 file:py-0.5 file:text-[9px] file:font-semibold file:text-white"
                    />
                  </div>

                  <div className="rounded-xl border border-dashed border-red-500/40 bg-red-950/20 p-2.5">
                    <h3 className="font-bold text-[10px] text-red-400 mb-1">📕 PDF (.pdf)</h3>
                    <input
                      type="file"
                      accept=".pdf"
                      onChange={handlePdfUpload}
                      className="w-full cursor-pointer rounded bg-slate-800 p-1 text-[9px] text-slate-300 file:mr-1 file:rounded file:border-0 file:bg-red-600 file:px-1.5 file:py-0.5 file:text-[9px] file:font-semibold file:text-white"
                    />
                  </div>
                </div>
              )}
            </div>

            {/* Question Selector Tabs */}
            <div className="mb-6 flex items-center justify-between border-b border-white/10 pb-4">
              <div className="flex max-h-36 overflow-y-auto flex-wrap gap-2 p-1">
                {mcqList.map((m, idx) => {
                  const isComplete = m.question && m.options.some((o) => o) && m.answer;
                  return (
                    <button
                      key={idx}
                      onClick={() => setCurrentIndex(idx)}
                      className={`rounded-lg px-3 py-1.5 text-xs font-bold transition flex items-center gap-1 ${
                        currentIndex === idx
                          ? "bg-blue-600 text-white ring-2 ring-blue-400"
                          : "bg-slate-800 text-slate-400 hover:bg-slate-700"
                      }`}
                    >
                      <span>Q{idx + 1}</span>
                      <span className={`h-2 w-2 rounded-full ${isComplete ? "bg-green-400" : "bg-red-500 animate-pulse"}`}></span>
                    </button>
                  );
                })}
              </div>

              <button
                onClick={handleAddQuestion}
                className="ml-2 shrink-0 rounded-lg bg-green-600 px-3 py-2 text-xs font-bold transition hover:bg-green-500"
              >
                ➕ Add Question
              </button>
            </div>

            <div className="flex items-center justify-between mb-4">
              <h2 className="text-xl font-bold text-slate-200">
                📝 Question {currentIndex + 1} Details
              </h2>
              {mcqList.length > 1 && (
                <button
                  onClick={() => handleDeleteQuestion(currentIndex)}
                  className="text-xs text-red-400 hover:text-red-300 underline font-semibold"
                >
                  🗑️ Delete Q{currentIndex + 1}
                </button>
              )}
            </div>

            {/* CUSTOM VOICE UPLOAD */}
            <div className="mb-4 rounded-xl border border-green-500/30 bg-green-950/20 p-3">
              <label className="block text-xs font-bold text-green-400 mb-1">
                🎙️ Upload Your Own Recorded Voice (.mp3 / .wav)
              </label>
              <input
                type="file"
                accept="audio/*"
                onChange={handleCustomAudioUpload}
                className="w-full cursor-pointer rounded bg-slate-800 p-1.5 text-xs text-slate-300 file:mr-2 file:rounded file:border-0 file:bg-green-600 file:px-2 file:py-1 file:text-xs file:font-semibold file:text-white"
              />
              {currentMCQ?.customAudioUrl && (
                <p className="mt-1 text-[10px] text-green-300 font-bold">
                  ✓ आपकी अपनी आवाज़ इस प्रश्न से जुड़ गई है।
                </p>
              )}
            </div>

            <div className="grid grid-cols-2 gap-3 mb-4">
              <div>
                <label className="mb-1 block text-xs font-semibold text-slate-300">Topic</label>
                <input
                  value={currentMCQ?.topic || ""}
                  onChange={(e) => updateCurrentMCQ("topic", e.target.value)}
                  className="w-full rounded-xl border border-white/10 bg-slate-800 px-3 py-2 text-xs outline-none focus:border-blue-500"
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-semibold text-slate-300">Exam / Tag</label>
                <input
                  value={currentMCQ?.tag || "GK Set"}
                  onChange={(e) => updateCurrentMCQ("tag", e.target.value)}
                  className="w-full rounded-xl border border-white/10 bg-slate-800 px-3 py-2 text-xs outline-none focus:border-blue-500"
                  placeholder="उदा: UP SI / UPSC / Easy"
                />
              </div>
            </div>

            <label className="mb-2 block text-sm font-semibold text-slate-300">Question</label>
            <textarea
              value={currentMCQ?.question || ""}
              onChange={(e) => updateCurrentMCQ("question", e.target.value)}
              className="mb-5 h-20 w-full rounded-xl border border-white/10 bg-slate-800 px-4 py-3 outline-none focus:border-blue-500"
              placeholder="यहाँ प्रश्न लिखें..."
            />

            <label className="mb-3 block text-sm font-semibold text-slate-300">Options</label>
            <div className="space-y-3">
              {currentMCQ?.options.map((option, index) => (
                <input
                  key={index}
                  value={option}
                  onChange={(e) => {
                    const newOpts = [...currentMCQ.options];
                    newOpts[index] = e.target.value;
                    updateCurrentMCQ("options", newOpts);
                  }}
                  className="w-full rounded-xl border border-white/10 bg-slate-800 px-4 py-3 outline-none focus:border-blue-500"
                  placeholder={`Option ${String.fromCharCode(65 + index)}`}
                />
              ))}
            </div>

            <label className="mb-2 mt-5 block text-sm font-semibold text-slate-300">Correct Answer</label>
            <select
              value={currentMCQ?.answer || ""}
              onChange={(e) => updateCurrentMCQ("answer", e.target.value)}
              className="w-full rounded-xl border border-white/10 bg-slate-800 px-4 py-3"
            >
              <option value="">सही उत्तर चुनें</option>
              {currentMCQ?.options.map((option, index) => (
                <option key={index} value={option}>
                  {String.fromCharCode(65 + index)}. {option}
                </option>
              ))}
            </select>

            <label className="mb-2 mt-5 block text-sm font-semibold text-slate-300">Explanation</label>
            <textarea
              value={currentMCQ?.explanation || ""}
              onChange={(e) => updateCurrentMCQ("explanation", e.target.value)}
              className="h-24 w-full rounded-xl border border-white/10 bg-slate-800 px-4 py-3 outline-none focus:border-blue-500"
              placeholder="यहाँ व्याख्या लिखें..."
            />

            <button
              onClick={handleDownloadPPT}
              className="mt-6 w-full rounded-xl bg-green-600 px-5 py-4 font-bold transition hover:bg-green-500 text-lg shadow-xl"
            >
              ⬇️ Export {activeTheme.name} Presentation (.pptx)
            </button>
          </section>

          {/* RIGHT PANEL PREVIEW WITH GEMINI VOICE ENGINE */}
          <section className="rounded-2xl border border-white/10 bg-slate-900 p-6 shadow-2xl">
            <div className="mb-4 flex items-center justify-between">
              <h2 className="text-xl font-bold">🎨 Live Preview & Gemini Audio</h2>
              <div className="flex gap-2">
                <button
                  disabled={currentIndex === 0}
                  onClick={() => setCurrentIndex(currentIndex - 1)}
                  className="rounded-lg bg-slate-800 px-3 py-1 text-xs text-slate-300 disabled:opacity-40"
                >
                  ◀ Prev
                </button>
                <button
                  disabled={currentIndex === mcqList.length - 1}
                  onClick={() => setCurrentIndex(currentIndex + 1)}
                  className="rounded-lg bg-slate-800 px-3 py-1 text-xs text-slate-300 disabled:opacity-40"
                >
                  Next ▶
                </button>
              </div>
            </div>

            {/* GEMINI VOICE CONTROLLER */}
            <div className="mb-5 flex items-center justify-between rounded-xl border border-blue-500/30 bg-blue-950/20 p-3.5">
              <div className="flex items-center gap-3">
                {!isSpeaking ? (
                  <button
                    onClick={handleGenerateGeminiVoice}
                    disabled={isGeneratingAudio}
                    className="rounded-lg bg-blue-600 px-4 py-2.5 text-xs font-bold transition hover:bg-blue-500 text-white shadow-md flex items-center gap-2 disabled:opacity-50"
                  >
                    {isGeneratingAudio ? "⏳ Processing..." : "🎙️ Play Gemini Voice"}
                  </button>
                ) : (
                  <button
                    onClick={handleStopVoice}
                    className="rounded-lg bg-red-600 px-4 py-2.5 text-xs font-bold transition hover:bg-red-500 text-white shadow-md flex items-center gap-2 animate-pulse"
                  >
                    ⏹️ Stop Audio
                  </button>
                )}
              </div>

              {currentMCQ?.customAudioUrl && (
                <audio src={currentMCQ.customAudioUrl} controls className="h-8 w-44" />
              )}
            </div>

            <div
              className="overflow-hidden rounded-2xl p-6 shadow-2xl min-h-[480px] flex flex-col justify-between border transition-all duration-300"
              style={{
                backgroundColor: `#${activeTheme.bg}`,
                borderColor: `#${activeTheme.border}`,
                fontFamily: selectedFont,
              }}
            >
              <div>
                <div
                  className="flex items-center justify-between mb-4 border-b pb-3 px-4 py-2 rounded-xl"
                  style={{ backgroundColor: `#${activeTheme.cardBg}` }}
                >
                  <span
                    className="font-bold text-sm"
                    style={{ color: `#${activeTheme.titleColor}` }}
                  >
                    विषय: {currentMCQ?.topic} {currentMCQ?.tag && `[${currentMCQ.tag}]`}
                  </span>
                  <span
                    className="font-bold text-xs"
                    style={{ color: `#${activeTheme.subTextColor}` }}
                  >
                    प्रश्न {currentIndex + 1} / {mcqList.length}
                  </span>
                </div>

                <div
                  className="mb-6 rounded-xl border p-5 shadow"
                  style={{
                    backgroundColor: `#${activeTheme.cardBg}`,
                    borderColor: `#${activeTheme.titleColor}55`,
                  }}
                >
                  <h3
                    className="font-bold"
                    style={{
                      color: `#${activeTheme.textColor}`,
                      fontSize: fontSize === "small" ? "1rem" : fontSize === "medium" ? "1.25rem" : "1.5rem",
                    }}
                  >
                    Q{currentIndex + 1}. {currentMCQ?.question || "यहाँ प्रश्न दिखेगा..."}
                  </h3>
                </div>

                <div className={layoutMode === "grid" ? "grid grid-cols-2 gap-3" : "space-y-2.5"}>
                  {currentMCQ?.options.map((option, index) => {
                    const isCorrect = option && option === currentMCQ.answer;
                    return (
                      <div
                        key={index}
                        className="rounded-xl px-4 py-3 font-semibold text-sm border shadow transition"
                        style={{
                          backgroundColor: isCorrect
                            ? `#${activeTheme.correctBg}`
                            : `#${activeTheme.cardBg}`,
                          borderColor: isCorrect
                            ? `#${activeTheme.correctColor}`
                            : `#${activeTheme.border}`,
                          color: isCorrect
                            ? `#${activeTheme.correctColor}`
                            : `#${activeTheme.textColor}`,
                        }}
                      >
                        ({String.fromCharCode(65 + index)}) {option || `Option ${String.fromCharCode(65 + index)}`}
                      </div>
                    );
                  })}
                </div>
              </div>

              {currentMCQ?.explanation && (
                <div
                  className="mt-6 rounded-xl border p-4"
                  style={{
                    backgroundColor: `#${activeTheme.correctBg}`,
                    borderColor: `#${activeTheme.correctColor}`,
                  }}
                >
                  <p
                    className="font-bold text-sm mb-1"
                    style={{ color: `#${activeTheme.correctColor}` }}
                  >
                    ✓ सही उत्तर: {currentMCQ.answer}
                  </p>
                  <p
                    className="text-xs leading-relaxed"
                    style={{ color: `#${activeTheme.textColor}` }}
                  >
                    <strong>व्याख्या:</strong> {currentMCQ.explanation}
                  </p>
                </div>
              )}

              {brandingName.trim() && (
                <div
                  className="mt-4 text-right text-xs font-bold opacity-75"
                  style={{ color: `#${activeTheme.subTextColor}` }}
                >
                  📢 {brandingName}
                </div>
              )}
            </div>
          </section>
        </div>
      )}
    </main>
  );
}