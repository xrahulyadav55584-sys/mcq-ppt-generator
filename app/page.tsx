"use client";

import { useState, useEffect, useRef } from "react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

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

const cleanText = (text: string) => {
  if (!text) return "";
  return text
    .replace(/[\*\_\`]/g, "")
    .replace(/^[🌍📖✅✔•\-\s]+/gu, "")
    .trim();
};

export default function Home() {
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

  const [fontSize, setFontSize] = useState<"small" | "medium" | "large">("medium");
  const [layoutMode, setLayoutMode] = useState<"grid" | "vertical">("grid");

  // Gemini API Voice States
  const [geminiApiKey, setGeminiApiKey] = useState("");
  const [isGeneratingAudio, setIsGeneratingAudio] = useState(false);
  const [isSpeaking, setIsSpeaking] = useState(false);
  const audioRef = useRef<HTMLAudioElement | null>(null);

  const currentMCQ = mcqList[currentIndex] || mcqList[0];
  const activeTheme = DESIGN_THEMES[selectedThemeKey];

  const fontSizes = {
    small: { question: 15, option: 14, title: 12 },
    medium: { question: 18, option: 16, title: 13 },
    large: { question: 22, option: 18, title: 14 },
  }[fontSize];

  useEffect(() => {
    if (typeof window !== "undefined" && "speechSynthesis" in window) {
      window.speechSynthesis.cancel();
      setIsSpeaking(false);
    }
  }, [currentIndex]);

  // Helper Speech Function
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

  // --- SAFE GEMINI API & TTS VOICE HANDLER ---
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

  // Upload Custom Voice File
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

  // --- SMART & ROBUST TEXT PARSER ---
  const parseRawTextToMCQs = (rawText: string): MCQ[] => {
    if (!rawText) return [];

    // शीर्षक से विषय पहचानने का प्रयास (उदा. 🌍 भूगोल — 20 महत्वपूर्ण प्रश्न)
    let detectedTopic = "सामान्य ज्ञान (GK)";
    const topicHeaderMatch = rawText.match(/^[^\n\d]*?([\u0900-\u097F\w\s]+?)(?:—|–|-|\d)/i);
    if (topicHeaderMatch && topicHeaderMatch[1].trim()) {
      const foundTopic = topicHeaderMatch[1].replace(/^[🌍📖✅✔•\-\s]+/gu, "").trim();
      if (foundTopic.length > 2) detectedTopic = foundTopic;
    }

    // प्रश्नों को अलग करने का रेगेक्स
    const blocks = rawText.split(/(?=\n?\s*(?:प्रश्न\s*\d+|\b\d+\b|Q\d+)[\.\:\-\)\s]+)/gi);
    const parsedMcqs: MCQ[] = [];

    blocks.forEach((block, idx) => {
      const lines = block
        .split("\n")
        .map((l) => l.trim())
        .filter((l) => l.length > 0);

      if (lines.length === 0) return;

      let qText = "";
      let rawOpts: { [key: string]: string } = {};
      let optsArray: string[] = [];
      let ansText = "";
      let expText = "";

      lines.forEach((line) => {
        // लाइन से इमोजी और स्पेशल सिंबल साफ करके अस्थायी रूप से चेक करें
        const cleanLine = line.replace(/^[🌍📖✅✔•\-\s]+/gu, "").trim();

        // 1. प्रश्न की लाइन
        if (!qText && cleanLine.match(/^(?:प्रश्न\s*\d+|\b\d+\b|Q\d+)[\.\:\-\)\s]+/i)) {
          qText = cleanLine.replace(/^(?:प्रश्न\s*\d+|\b\d+\b|Q\d+)[\.\:\-\)\s]+\s*/i, "").trim();
        } 
        // 2. विकल्प A, B, C, D
        else if (cleanLine.match(/^[A-D][\.\:\-\)\s]+/i)) {
          const letter = cleanLine.charAt(0).toUpperCase();
          const optContent = cleanLine.replace(/^[A-D][\.\:\-\)\s]+\s*/i, "").trim();
          rawOpts[letter] = cleanText(optContent);
        } 
        // 3. उत्तर पहचानें
        else if (cleanLine.match(/(?:सही उत्तर|उत्तर|Answer|Ans)[\:\-\s]+/i)) {
          let ansValue = cleanLine.replace(/.*?(?:सही उत्तर|उत्तर|Answer|Ans)[\:\-\s]+\s*/i, "").trim();
          const optMatch = ansValue.match(/^([A-D])[\.\:\-\)\s]*/i);
          if (optMatch) {
            const letter = optMatch[1].toUpperCase();
            ansValue = rawOpts[letter] || ansValue.replace(/^[A-D][\.\:\-\)\s]+\s*/i, "");
          } else {
            ansValue = ansValue.replace(/^[A-D][\.\:\-\)\s]+\s*/i, "");
          }
          ansText = cleanText(ansValue);
        } 
        // 4. व्याख्या पहचानें
        else if (cleanLine.match(/(?:व्याख्या|Explanation|Exp)[\:\-\s]+/i)) {
          expText = cleanLine.replace(/.*?(?:व्याख्या|Explanation|Exp)[\:\-\s]+\s*/i, "").trim();
        }
      });

      if (Object.keys(rawOpts).length > 0) {
        optsArray = [
          rawOpts["A"] || "",
          rawOpts["B"] || "",
          rawOpts["C"] || "",
          rawOpts["D"] || "",
        ];
      }

      if (qText && optsArray.some((o) => o !== "")) {
        let matchedAns = ansText;

        if (!optsArray.includes(ansText)) {
          const found = optsArray.find((o) => o.toLowerCase() === ansText.toLowerCase());
          if (found) matchedAns = found;
          else if (optsArray[0]) matchedAns = optsArray[0];
        }

        parsedMcqs.push({
          id: Date.now() + idx,
          topic: detectedTopic,
          question: cleanText(qText),
          options: optsArray,
          answer: matchedAns,
          explanation: cleanText(expText) || "व्याख्या उपलब्ध नहीं है।",
          tag: "GK Set",
        });
      }
    });

    return parsedMcqs;
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
        const pageText = textContent.items
          .map((item: any) => item.str)
          .join("\n");
        fullText += pageText + "\n";
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

  // PPT Export Engine
  const handleDownloadPPT = async () => {
    const validMcqs = mcqList.filter((m) => m.question.trim() !== "");

    if (validMcqs.length === 0) {
      alert("कृपया कम से कम एक प्रश्न दर्ज करें!");
      return;
    }

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";
    const theme = DESIGN_THEMES[selectedThemeKey];

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
      color: theme.titleColor,
      align: "left",
    });

    coverSlide.addText("बहुविकल्पीय प्रश्नोत्तरी एवं विस्तृत व्याख्या सेट", {
      x: 1.2,
      y: 2.3,
      w: 7.6,
      h: 0.5,
      fontSize: 18,
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
        fontSize: fontSizes.title,
        color: theme.titleColor,
        bold: true,
      });

      slide1.addText(`प्रश्न ${i + 1} / ${validMcqs.length}`, {
        x: 6.6,
        y: 0.4,
        w: 2.4,
        h: 0.5,
        fontSize: fontSizes.title,
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
        fontSize: fontSizes.question,
        bold: true,
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
            fontSize: fontSizes.option,
            bold: true,
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
            fontSize: fontSizes.option,
            bold: true,
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
        fontSize: fontSizes.title,
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
        fontSize: fontSizes.question,
        bold: true,
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
        color: theme.titleColor,
      });

      slide2.addText(mcq.explanation, {
        x: 1.1,
        y: 2.9,
        w: 7.8,
        h: 2.0,
        fontSize: fontSizes.option,
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
      color: theme.titleColor,
    });

    const rows: any[] = [
      [
        { text: "क्र.", options: { bold: true, fill: theme.cardBg, color: theme.titleColor } },
        { text: "प्रश्न (Question)", options: { bold: true, fill: theme.cardBg, color: theme.titleColor } },
        { text: "सही उत्तर", options: { bold: true, fill: theme.cardBg, color: theme.correctColor } },
      ],
    ];

    validMcqs.forEach((mcq, idx) => {
      rows.push([
        { text: `Q${idx + 1}`, options: { bold: true, color: theme.textColor } },
        { text: mcq.question.slice(0, 45) + "...", options: { color: theme.subTextColor } },
        { text: mcq.answer, options: { bold: true, color: theme.correctColor } },
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
      <header className="border-b border-white/10 bg-slate-900">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-5">
          <div>
            <h1 className="text-2xl font-bold">MCQ → Gamma PPT & Gemini Voice Suite</h1>
            <p className="mt-1 text-sm text-slate-400">
              Free Gemini API Key Supported Presentation & Voice Tool
            </p>
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={handleCopyTextPaper}
              className="rounded-lg bg-blue-600/30 px-3 py-2 text-xs font-bold text-blue-400 border border-blue-500/40 hover:bg-blue-600/50"
            >
              📋 Copy Text Paper
            </button>
            <button
              onClick={handleClearAll}
              className="rounded-lg bg-red-600/20 px-3 py-2 text-xs font-bold text-red-400 border border-red-500/40 hover:bg-red-600/40"
            >
              🧹 Clear All
            </button>
          </div>
        </div>
      </header>

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
    </main>
  );
}