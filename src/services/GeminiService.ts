const SYSTEM_INSTRUCTION = `Bạn là chuyên gia Excel 365. PHẢI TRẢ VỀ JSON theo schema sau:
{
  "type": "single" | "array" | "chart",
  "value": "nếu là single (ví dụ công thức =SUM...)",
  "values": ["nếu là array (danh sách công thức/giá trị)"],
  "chartData": {
    "type": "pie" | "column" | "line",
    "title": "Tiêu đề",
    "range": "Vùng dữ liệu",
    "table": [["Tiêu đề 1", "Tiêu đề 2"], ["Dòng 1 cột 1", "Dòng 1 cột 2"]]
  }
}

QUY TẮC:
1. SMART CHART: Nếu vùng yêu cầu (vd A21:B25) chỉ có chữ, hãy tự tạo "table" tổng hợp đếm số lượng.
2. Filter: Luôn thêm (vùng_điều_kiện<>"") để bỏ qua ô trống.
3. Tuyệt đối KHÔNG GIẢI THÍCH bản văn. Chỉ JSON.`;

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
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;

  const { data, usedRangeAddress, activeCellAddress, allSheetNames } = excelContext;

  const body = {
    contents: [
      {
        parts: [
          {
            text: `${SYSTEM_INSTRUCTION}
            
DỮ LIỆU CONTEXT:
- Vùng dữ liệu đang dùng (Used Range): ${usedRangeAddress}
- Ô đang chọn (Active Cell): ${activeCellAddress}
- Danh sách tất cả các Sheet: ${allSheetNames?.join(', ')}
- Dữ liệu mẫu (JSON): ${JSON.stringify(data)}

${intent ? `[INTENT: ${intent}]` : ''}
YÊU CẦU: ${prompt}`
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
