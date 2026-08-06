const { GoogleGenAI } = require('@google/genai');
const PptxGenJS = require('pptxgenjs');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { prompt, answers } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: 'Thiếu thông tin prompt!' });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: 'Chưa cấu hình GEMINI_API_KEY trên Vercel!' });
    }

    const ai = new GoogleGenAI({ apiKey });

    const systemPrompt = `
      Bạn là một chuyên gia thiết kế bài thuyết trình. Hãy tạo nội dung slide dựa trên ý tưởng: "${prompt}".
      Yêu cầu khán giả: ${answers?.q1 || 'Mọi người'}.
      Phong cách: ${answers?.q2 || 'Hiện đại'}.
      Mục tiêu: ${answers?.q3 || 'Tổng quát'}.

      Hãy trả về duy nhất một cấu trúc JSON hợp lệ (KHÔNG dùng markdown backticks, KHÔNG kèm lời dẫn) theo cấu trúc:
      {
        "title": "Tên bài thuyết trình",
        "slides": [
          {
            "slideTitle": "Tiêu đề Slide 1",
            "bullets": ["Ý chính 1", "Ý chính 2", "Ý chính 3"]
          }
        ]
      }
      Tạo khoảng 4 đến 5 slide nội dung chất lượng.
    `;

    // Sử dụng model mới nhất gemini-3.5-flash
    const response = await ai.models.generateContent({
      model: 'gemini-3.5-flash',
      contents: systemPrompt,
    });

    const responseText = response.text;
    const cleanJsonString = responseText.replace(/```json/gi, '').replace(/```/g, '').trim();
    const presentationData = JSON.parse(cleanJsonString);

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: '0F172A' };
    titleSlide.addText(presentationData.title || prompt, {
      x: 1, y: 2.2, w: '80%', h: 1.5,
      fontSize: 32, bold: true, color: 'FFFFFF', align: 'center'
    });

    if (presentationData.slides && Array.isArray(presentationData.slides)) {
      presentationData.slides.forEach((item) => {
        const slide = pptx.addSlide();
        slide.background = { color: 'F8FAFC' };

        slide.addText(item.slideTitle, {
          x: 0.8, y: 0.6, w: '85%', h: 0.8,
          fontSize: 24, bold: true, color: '1E293B'
        });

        const bulletText = item.bullets.map(b => ({ text: b, options: { bullet: true, breakLine: true } }));
        slide.addText(bulletText, {
          x: 0.8, y: 1.6, w: '85%', h: 4.5,
          fontSize: 16, color: '334155', lineSpacing: 28
        });
      });
    }

    const buffer = await pptx.write({ outputType: 'nodebuffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename=Slide_${Date.now()}.pptx`);
    return res.status(200).send(buffer);

  } catch (error) {
    console.error('Lỗi Backend:', error);
    return res.status(500).json({ error: error.message || 'Lỗi xử lý hệ thống!' });
  }
};