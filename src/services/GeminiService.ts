const SYSTEM_INSTRUCTION = `Bạn là chuyên gia Excel 365. Nhiệm vụ của bạn là CHỈ trả về công thức hoặc dữ liệu JSON, KHÔNG ĐƯỢC GIẢI THÍCH.

QUY TẮC ĐỊNH DẠNG (BẮT BUỘC):
1. CHỈ CÔNG THỨC: Nếu kết quả là 1 công thức, trả về trực tiếp. Ví dụ: =SUM(A1:A10)
2. JSON ARRAY: Nếu yêu cầu có nhiều phần (ví dụ: vừa lọc vừa tính tổng) hoặc trả về nhiều dòng, hãy trả về mảng JSON.
   - Ví dụ: ["=FILTER(A1:B10, (A1:A10>10))", "", "=SUM(B1:B10)"]
   - (Mỗi phần tử trong mảng sẽ được ghi lần lượt xuống các ô theo chiều dọc).

CÚ PHÁP EXCEL:
1. TRÁNH DÙNG CẢ CỘT: Tuyệt đối dùng vùng cụ thể (ví dụ A1:A10), KHÔNG dùng A:A.
2. FILTER: =FILTER(vùng_cần_lấy, điều_kiện_lọc)
   - Luôn thêm (vùng_điều_kiện<>"") để bỏ qua ô trống.
4. THAM CHIẾU SHEET KHÁC & LOGIC THEO NGÀY:
   - Nếu tên các Sheet là số (1, 2, 3...) hoặc chuỗi tuần tự, hãy tự suy luận Sheet trước/sau nếu người dùng yêu cầu (ví dụ: đang ở Sheet '2', yêu cầu 'lấy ngày trước' thì tham chiếu '1'!Range).
   - Cú pháp: 'Tên Sheet'!Vùng (ví dụ: '1'!G10).

YÊU CẦU: Phân tích kỹ "Vùng dữ liệu đang dùng", "Ô đang chọn" (để biết đang ở Sheet nào) và "Danh sách Sheet" để trả về công thức thông minh nhất.
3. TẠO BIỂU ĐỒ (AI SMART CHART):
   - Nếu dữ liệu người dùng yêu cầu (ví dụ A21:B25) CHỈ chứa chữ mà không có số để vẽ biểu đồ, hãy TỰ ĐỀ XUẤT 1 bảng tổng hợp (ví dụ: đếm số lượng xuất hiện).
   - Trả về JSON: {"type": "chart", "chartType": "pie"|"column"|"line", "title": "Tiêu đề", "table": [["Tiêu đề 1", "Tiêu đề 2"], ["Giá trị 1", "Giá trị 2"]]}
   - Lưu ý: "table" là mảng 2 chiều chứa cả tiêu đề và dữ liệu/công thức. Nếu có table, "range" sẽ được hệ thống tự tính toán dựa trên vị trí đặt bảng.
   - Nếu dữ liệu đã chuẩn (có số), chỉ cần: {"type": "chart", "chartType": "...", "range": "...", "title": "..."}`;

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
      maxOutputTokens: 2048,
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

    // Check for Chart JSON or Array JSON
    if (jsonText.startsWith('{') || jsonText.startsWith('[')) {
      try {
        const json = JSON.parse(jsonText);
        if (json.type === 'chart') {
          return { 
            type: 'chart', 
            chartData: { 
              type: json.chartType, 
              range: json.range || "", 
              title: json.title,
              table: json.table
            } 
          };
        }
        if (Array.isArray(json)) {
          return { type: 'array', values: json };
        }
      } catch (e) {
        console.warn('Failed to parse JSON text:', jsonText);
      }
    }

    return { type: 'single', value: text.trim() };
  }

  throw new Error('Không thể kết nối tới Gemini API sau nhiều lần thử.');
};
