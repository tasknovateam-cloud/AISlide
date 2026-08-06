import { GoogleGenAI } from '@google/genai';
import PptxGenJS from 'pptxgenjs';

export default async function handler(req, res) {
  // Chỉ chấp nhận phương thức POST từ Frontend
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  try {
    const { prompt, answers } = req.body;

    // 1. Khởi tạo Gemini AI (sử dụng API Key lưu trong môi trường)
    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

    // 2. Viết câu lệnh (Prompt) yêu cầu AI trả về định dạng JSON chuẩn
    const aiPrompt = `
      Bạn là chuyên gia thiết kế nội dung bài thuyết trình.
      Hãy tạo nội dung cho bài thuyết trình dựa trên ý tưởng: "${prompt}".
      Chi tiết bổ sung từ người dùng:
      - Khán giả: ${answers.q1 || 'Chung'}
      - Phong cách: ${answers.q2 || 'Hiện đại'}
      - Mục tiêu: ${answers.q3 || 'Tổng quan'}

      YÊU CẦU:
      Trả về kết quả duy nhất ở dạng chuỗi JSON thuần túy (không chứa markdown, không chứa \`\`\`json), là một mảng danh sách từ 4 đến 6 slide theo cấu trúc:
      [
        {
          "title": "Tiêu đề trang slide",
          "bullets": ["Ý chính 1", "Ý chính 2", "Ý chính 3"]
        }
      ]
    `;

    // 3. Gọi Gemini sinh nội dung
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: aiPrompt,
    });

    let rawText = response.text.trim();
    // Làm sạch chuỗi nếu AI lỡ bọc trong khối code ```json
    if (rawText.startsWith('```')) {
      rawText = rawText.replace(/^```json\n?/, '').replace(/```$/, '').trim();
    }

    const slidesData = JSON.parse(rawText);

    // 4. Khởi tạo file PowerPoint từ dữ liệu AI trả về
    const pptx = new PptxGenJS();
    
    // Đặt kích thước slide chuẩn 16:9
    pptx.layout = 'LAYOUT_16x9';

    slidesData.forEach((slideContent, index) => {
      const slide = pptx.addSlide();

      // Tô màu nền xám đen hiện đại cho Slide
      slide.background = { color: "0F172A" };

      // Tiêu đề Slide
      slide.addText(slideContent.title, {
        x: 0.8,
        y: 0.8,
        w: 8.4,
        h: 1.0,
        fontSize: 28,
        bold: true,
        color: "A78BFA", // Màu tím neon
        fontFace: "Arial"
      });

      // Các ý chính dạng Bullet Point
      if (slideContent.bullets && slideContent.bullets.length > 0) {
        const bulletText = slideContent.bullets.map(b => ({ text: b, options: { bullet: true, breakLine: true } }));
        slide.addText(bulletText, {
          x: 0.8,
          y: 2.0,
          w: 8.4,
          h: 4.5,
          fontSize: 18,
          color: "F8FAFC",
          lineSpacing: 32,
          fontFace: "Arial"
        });
      }
    });

    // 5. Xuất file dưới dạng Buffer và gửi về trình duyệt để tải xuống
    const buffer = await pptx.write({ outputType: 'nodebuffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="AI_Presentation.pptx"');
    return res.status(200).send(buffer);

  } catch (error) {
    console.error("Lỗi xử lý:", error);
    return res.status(500).json({ error: 'Có lỗi xảy ra khi tạo Slide: ' + error.message });
  }
}