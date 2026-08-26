function generatePPTX() {
    // PptxGenJS ob'ektini yaratish
    let pptx = new PptxGenJS();

    // Kiritilgan ma'lumotlarni olish
    let inputs = document.querySelectorAll('input');
    let topic = inputs[0] && inputs[0].value.trim() !== "" ? inputs[0].value : "Taqdimot";
    let author = inputs[1] && inputs[1].value.trim() !== "" ? inputs[1].value : "Muallif";

    // 16:9 slayd formati
    pptx.layout = 'LAYOUT_169';

    // 1-SLAYD: Muqova
    let slide1 = pptx.addSlide();
    slide1.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '0F172A' } });

    slide1.addText(topic, {
        x: 0.8, y: 2.2, w: 11.7, h: 1.8,
        fontSize: 36, bold: true, color: 'FFFFFF',
        align: 'center', fontFace: 'Calibri'
    });

    slide1.addText("Tayyorladi: " + author, {
        x: 0.8, y: 4.2, w: 11.7, h: 0.8,
        fontSize: 22, color: '38BDF8',
        align: 'center', fontFace: 'Calibri'
    });

    // 2-SLAYD: Asosiy mazmun
    let slide2 = pptx.addSlide();
    slide2.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '1E293B' } });

    slide2.addText(topic + " — Asosiy ma'lumotlar", {
        x: 0.8, y: 0.8, w: 11.7, h: 1.0,
        fontSize: 28, bold: true, color: 'FFFFFF', fontFace: 'Calibri'
    });

    slide2.addText("• Ketma-ket ulashda to'liq qarshilik: R = R1 + R2 + ... + Rn\n• Parallel ulashda to'liq qarshilik: 1/R = 1/R1 + 1/R2 + ... + 1/Rn\n• Ketma-ket ulashda tok kuchi barcha qarshiliklarda bir xil bo'ladi (I = I1 = I2).\n• Parallel ulashda kuchlanish barcha tarmoqlarda bir xil bo'ladi (U = U1 = U2).", {
        x: 0.8, y: 2.0, w: 11.7, h: 4.5,
        fontSize: 20, color: 'E2E8F0', lineSpacing: 32, fontFace: 'Calibri'
    });

    // Faylni kompyuterga saqlash
    pptx.writeFile({ fileName: `${topic}.pptx` });
}
