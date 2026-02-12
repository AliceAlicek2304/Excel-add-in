const SYSTEM_INSTRUCTION = `You are an Excel 365 expert. You MUST RETURN JSON according to the following schema:
{
  "type": "single" | "array" | "chart",
  "value": "if single (e.g., formula =SUM...)",
  "values": ["if array (list of formulas/values)"],
  "chartData": {
    "type": "pie" | "column" | "line",
    "title": "Title",
    "range": "Data range",
    "table": [["Header 1", "Header 2"], ["Row 1 Col 1", "Row 1 Col 2"]]
  }
}

RULES:
1. SMART CHART: If requested range (e.g., A21:B25) contains ONLY text, automatically generate a summary "table" (e.g., counting occurrences).
2. CONSOLIDATE: If intent is CONSOLIDATE_AND_CHART, create a "table" using FORMULAS referencing other sheets.
   - Use the provided "allSheetNames" list.
   - Example table: [["Sheet", "Value"], ["Sheet1", "='Sheet1'!F1"], ["Sheet2", "='Sheet2'!F1"]]
3. Filter: Always add (condition_range<>"") to ignore blank cells.
4. Absolutely NO text explanation. ONLY JSON.`;

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export interface GeminiResult {
  type: 'single' | 'array' | 'chart';
  value?: string;
  values?: string[];
  chartData?: {
    type: 'pie' | 'column' | 'line';
    range: string;
    title: string;
    // New fields for smart charting
    table?: any[][]; 
  };
}

export const processWithGemini = async (
  apiKey: string, 
  prompt: string, 
  excelContext: any,
  intent?: string
): Promise<GeminiResult> => {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${apiKey}`;

  const { data, usedRangeAddress, activeCellAddress, allSheetNames } = excelContext;

  const body = {
    contents: [
      {
        parts: [
          {
            text: `${SYSTEM_INSTRUCTION}
            
CONTEXT DATA:
- Used Range: ${usedRangeAddress}
- Active Cell: ${activeCellAddress}
- All Sheet Names: ${allSheetNames?.join(', ')}
- Sample Data (JSON): ${JSON.stringify(data)}

${intent ? `[INTENT: ${intent}]` : ''}
REQUEST: ${prompt}`
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 4096,
      response_mime_type: "application/json"
    }
  };

  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    console.log(`Attempt ${attempt}/${maxRetries}...`);

    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const errorData = await res.json().catch(() => ({}));
      if ((res.status === 404 || res.status === 429) && attempt < maxRetries) {
        const waitTime = attempt * 3000;
        await sleep(waitTime);
        continue;
      }
      throw new Error(errorData?.error?.message || `HTTP ${res.status}: ${res.statusText}`);
    }

    const resJson = await res.json();
    const text = resJson?.candidates?.[0]?.content?.parts?.[0]?.text;
    
    if (!text) {
      throw new Error('Không nhận được phản hồi từ AI.');
    }

    let jsonText = text.trim();

    // Attempt to extract JSON from markdown code blocks if they exist
    const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      jsonText = jsonMatch[1].trim();
    } else {
      // Fallback: search for first { or [ and last } or ]
      const firstBrace = text.indexOf('{');
      const firstBracket = text.indexOf('[');
      const lastBrace = text.lastIndexOf('}');
      const lastBracket = text.lastIndexOf(']');

      const start = (firstBrace !== -1 && (firstBracket === -1 || firstBrace < firstBracket)) ? firstBrace : firstBracket;
      const end = (lastBrace !== -1 && (lastBracket === -1 || lastBrace > lastBracket)) ? lastBrace : lastBracket;

      if (start !== -1 && end !== -1 && end > start) {
        jsonText = text.substring(start, end + 1).trim();
      }
    }

    try {
      const json = JSON.parse(jsonText);
      
      // Normalize response from JSON mode
      if (json.type === 'chart') {
        return { 
          type: 'chart', 
          chartData: { 
            type: json.chartData?.type || json.chartType || 'column', 
            range: json.chartData?.range || json.range || "", 
            title: json.chartData?.title || json.title || "Biểu đồ AI",
            table: json.chartData?.table || json.table
          } 
        };
      }
      
      if (json.type === 'array' || Array.isArray(json.values)) {
        return { type: 'array', values: json.values };
      }

      if (json.type === 'single' || json.value) {
        return { type: 'single', value: json.value || json.toString() };
      }

      // Final fallback for direct structure
      return json as GeminiResult;
    } catch (e) {
      console.warn('Failed to parse JSON text:', jsonText);
      return { type: 'single', value: text.trim() };
    }
  }

  throw new Error('Không thể kết nối tới Gemini API sau nhiều lần thử.');
};
