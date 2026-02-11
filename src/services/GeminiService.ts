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
3. KHÔNG GIẢI THÍCH: Tuyệt đối không thêm bất kỳ văn bản nào như "Dưới đây là công thức..." hay "Ghi chú:". Nếu vi phạm, kết quả sẽ bị lỗi.

YÊU CẦU: Phân tích kỹ "Vùng dữ liệu đang dùng" để chọn địa chỉ ô chính xác.`;

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export interface GeminiResult {
  type: 'single' | 'array';
  value?: string;
  values?: string[];
}

export const processWithGemini = async (apiKey: string, prompt: string, excelContext: any): Promise<GeminiResult> => {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;

  const { data, usedRangeAddress, activeCellAddress } = excelContext;

  const body = {
    contents: [
      {
        parts: [
          {
            text: `${SYSTEM_INSTRUCTION}
            
DỮ LIỆU CONTEXT:
- Vùng dữ liệu đang dùng (Used Range): ${usedRangeAddress}
- Ô đang chọn (Active Cell): ${activeCellAddress}
- Dữ liệu mẫu (JSON): ${JSON.stringify(data)}

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

    const trimmed = text.trim();

    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      try {
        const array = JSON.parse(trimmed);
        if (Array.isArray(array)) {
          return { type: 'array', values: array };
        }
      } catch (e) {
        console.warn('Failed to parse array, treating as single value');
      }
    }

    return { type: 'single', value: trimmed };
  }

  throw new Error('Không thể kết nối tới Gemini API sau nhiều lần thử.');
};
