const SYSTEM_INSTRUCTION = `Bạn là chuyên gia Excel 365. Phân tích yêu cầu và trả về công thức Excel CHÍNH XÁC.

CÚ PHÁP QUAN TRỌNG:
1. TRÁNH DÙNG CẢ CỘT: Tuyệt đối KHÔNG dùng các vùng như A:A, B:B để tránh lỗi #SPILL! và tăng tốc độ.
   - Hãy sử dụng vùng cụ thể dựa trên thông tin "Vùng dữ liệu đang dùng" (Used Range) được cung cấp.
   - Ví dụ: Thay vì FILTER(A:A,...), hãy dùng FILTER(A1:A100,...).

2. FILTER: =FILTER(vùng_cần_lấy, điều_kiện_lọc)
   - Luôn thêm điều kiện loại bỏ ô trống: (vùng_điều_kiện<>"")
   - Ví dụ lọc cột A:B khi A > 10: =FILTER(A1:B100, (A1:A100>10)*(A1:A100<>""))

3. VLOOKUP: =VLOOKUP(giá_trị_tìm, vùng_tìm, số_cột, FALSE)
   - Vùng tìm nên là vùng cụ thể (ví dụ A1:B100).

4. Vùng ô: Dùng dấu hai chấm (:).

QUAN TRỌNG:
- Dữ liệu JSON được cung cấp là mẫu 50 dòng đầu tiên.
- "Vùng dữ liệu đang dùng" (Used Range) cho biết giới hạn thực tế của bảng.
- Trả về TRỰC TIẾP công thức (bắt đầu bằng =) hoặc JSON array ["giá trị 1","giá trị 2"] nếu cần ghi vào nhiều ô.
- CHỈ trả về kết quả, KHÔNG giải thích.`;

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
