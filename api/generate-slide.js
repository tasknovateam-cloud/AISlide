import { GoogleGenerativeAI } from "@google/generative-ai";
import pptxgen from "pptxgenjs";

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

export default async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ message: 'Method not allowed' });
    }

    const { action, topic, answers } = req.body;

    if (!topic) {
        return res.status(400).json({ message: 'Vui lòng cung cấp chủ đề.' });
    }

    try {
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

        // BƯỚC 1: LẤY 3 CÂU HỎI CỐT LÕI DỰA TRÊN CHỦ ĐỀ
        if (action === 'get-questions') {
            const promptQuestions = `Bạn là chuyên gia thuyết trình. Hãy đưa ra 3 câu hỏi cốt lõi nhất về chủ đề "${topic}" để giúp người dùng làm rõ mục tiêu, đối tượng hướng đến hoặc nội dung quan trọng nhất cho bài thuyết trình. 
            Trả về kết quả dưới dạng JSON duy nhất có cấu trúc:
            {
              "questions": [
                "Câu hỏi 1...",
                "Câu hỏi 2...",
                "Câu hỏi 3..."
              ]
            }`;

            const result = await model.generateContent(promptQuestions);
            let responseText = result.response.text().trim();
            responseText = responseText.replace(/```json/g, '').replace(/```/g, '');
            const data = JSON.parse(responseText);

            return res.status(200).json(data);
        }

        // BƯỚC 2: TẠO FILE PPTX TỪ CHỦ ĐỀ VÀ 3 CÂU TRẢ LỜI
        if (action === 'generate-pptx') {
            const promptSlide = `Tạo nội dung bài thuyết trình PowerPoint về chủ đề: "${topic}".
            Người dùng đã trả lời 3 câu hỏi định hướng như sau:
            ${answers ? answers.map((a, i) => `${i + 1}. ${a.question} -> Trả lời: ${a.answer}`).join('\n') : ''}

            Hãy tạo bài thuyết trình gồm 3-5 slide. Trả về định dạng JSON thuần túy theo cấu trúc:
            {
              "slides": [
                {
                  "title": "Tiêu đề slide",
                  "subtitle": "Mô tả ngắn hoặc phụ đề",
                  "bullets": ["Ý chính 1", "Ý chính 2", "Ý chính 3"],
                  "speakerNotes": "Ghi chú cho diễn giả khi trình bày slide này"
                }
              ]
            }`;

            const result = await model.generateContent(promptSlide);
            let responseText = result.response.text().trim();
            responseText = responseText.replace(/```json/g, '').replace(/```/g, '');
            const slideData = JSON.parse(responseText);

            // Dùng pptxgenjs đóng gói file .pptx
            const pptx = new pptxgen();
            pptx.layout = 'LAYOUT_16x9';

            slideData.slides.forEach((slideContent) => {
                const slide = pptx.addSlide();
                
                // Thiết lập background tối sang trọng
                slide.background = { color: "0F172A" };

                // Tiêu đề
                slide.addText(slideContent.title, {
                    x: 0.8, y: 0.6, w: '85%', h: 1.0,
                    fontSize: 28, bold: true, color: "F8FAFC", fontFace: "Segoe UI"
                });

                // Phụ đề
                if (slideContent.subtitle) {
                    slide.addText(slideContent.subtitle, {
                        x: 0.8, y: 1.5, w: '85%', h: 0.6,
                        fontSize: 16, color: "94A3B8", italic: true, fontFace: "Segoe UI"
                    });
                }

                // Bullets nội dung
                if (slideContent.bullets && slideContent.bullets.length > 0) {
                    const bulletText = slideContent.bullets.map(b => ({ text: b, options: { bullet: true, breakLine: true } }));
                    slide.addText(bulletText, {
                        x: 0.8, y: 2.3, w: '85%', h: 4.0,
                        fontSize: 18, color: "CBD5E1", lineSpacing: 28, fontFace: "Segoe UI"
                    });
                }

                // Ghi chú thuyết trình
                if (slideContent.speakerNotes) {
                    slide.addNotes(slideContent.speakerNotes);
                }
            });

            // Xuất file dưới dạng Buffer để gửi về Client download
            const buffer = await pptx.write({ outputType: "nodebuffer" });

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
            res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(topic)}.pptx"`);
            return res.status(200).send(buffer);
        }

    } catch (error) {
        console.error("API Error:", error);
        return res.status(500).json({ message: "Đã xảy ra lỗi: " + error.message });
    }
}