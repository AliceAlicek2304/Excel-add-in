const SYSTEM_INSTRUCTION = `Bạn là chuyên gia Excel 365. Phân tích yêu cầu và trả về công thức Excel CHÍNH XÁC.

CÚ PHÁP QUAN TRỌNG:
1. FILTER: =FILTER(vùng_cần_lấy, điều_kiện_lọc)
   - Lọc cột A và B khi A < 10000 (BỎ QUA Ô TRỐNG): =FILTER(A:B,(A:A<10000)*(A:A<>""))
   - Lọc nhiều điều kiện: =FILTER(A:B,(A:A>100)*(A:A<1000)*(A:A<>""))
   - QUAN TRỌNG: Luôn thêm *(A:A<>"") để loại bỏ ô trống khi lọc toàn bộ cột

2. VLOOKUP: =VLOOKUP(giá_trị_tìm,vùng_tìm,số_cột,FALSE)

3. Vùng ô: Dùng dấu hai chấm (:) - Ví dụ: A1:A10, B:B

QUAN TRỌNG: Dữ liệu JSON có key là chữ cái (A, B, C...) tương ứng với tên cột Excel.

Trả về:
- Nếu 1 công thức/giá trị: Trả về trực tiếp (bắt đầu = nếu là công thức)
- Nếu nhiều giá trị: Trả về JSON array ["giá trị 1","giá trị 2"]

CHỈ trả về kết quả, KHÔNG giải thích.`;

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export interface GeminiResult {
  type: 'single' | 'array';
  value?: string;
  values?: string[];
}

export const processWithGemini = async (apiKey: string, prompt: string, jsonData: any): Promise<GeminiResult> => {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;

  const body = {
    contents: [
      {
        parts: [
          {
            text: `${SYSTEM_INSTRUCTION}\n\nDữ liệu Excel (JSON): ${JSON.stringify(jsonData)}\n\nYêu cầu: ${prompt}`
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

    console.log('Response status:', res.status, res.statusText);

    if (!res.ok) {
      const errorData = await res.json().catch(() => ({}));
      console.error('API Error:', JSON.stringify(errorData, null, 2));

      // Retry on 404 or 429 (rate limit)
      if ((res.status === 404 || res.status === 429) && attempt < maxRetries) {
        const waitTime = attempt * 3000; // 3s, 6s
        console.log(`Retrying in ${waitTime / 1000}s...`);
        await sleep(waitTime);
        continue;
      }

      throw new Error(errorData?.error?.message || `HTTP ${res.status}: ${res.statusText}`);
    }

    const data = await res.json();

    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) {
      throw new Error('Không nhận được phản hồi từ AI.');
    }

    const trimmed = text.trim();

    // Check if result is a JSON array (multiple values)
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

    // Single value
    return { type: 'single', value: trimmed };
  }

  throw new Error('Không thể kết nối tới Gemini API sau nhiều lần thử. Vui lòng đợi vài giây rồi thử lại.');
};

export interface GeminiResult {
  type: 'single' | 'array';
  value?: string;
  values?: string[];
}
