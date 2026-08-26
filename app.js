function generatePPTX() {
    let pptx = new PptxGenJS();

    let inputs = document.querySelectorAll('input');
    let topic = inputs[0] && inputs[0].value.trim() !== "" ? inputs[0].value : "Taqdimot";
    let author = inputs[1] && inputs[1].value.trim() !== "" ? inputs[1].value : "Muallif";

    // 1-SLAYD: Muqova
    let slide1 = pptx.addSlide();
    slide1.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '0F172A' } });

    slide1.addText(topic, {
        x: 0.8, y: 2.2, w: 8.4, h: 1.8,
        fontSize: 32, bold: true, color: 'FFFFFF',
        align: 'center', fontFace: 'Calibri'
    });

    slide1.addText("Tayyorladi: " + author, {
        x: 0.8, y: 4.2, w: 8.4, h: 0.8,
        fontSize: 20, color: '38BDF8',
        align: 'center', fontFace: 'Calibri'
    });

    // 2-SLAYD: Asosiy mazmun
    let slide2 = pptx.addSlide();
    slide2.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '1E293B' } });

    slide2.addText(topic + " — Asosiy ma'lumotlar", {
        x: 0.8, y: 0.8, w: 8.4, h: 1.0,
        fontSize: 24, bold: true, color: 'FFFFFF', fontFace: 'Calibri'
    });

    slide2.addText("• Ketma-ket ulashda sig'im: 1/C = 1/C1 + 1/C2 + ... + 1/Cn\n• Parallel ulashda sig'im: C = C1 + C2 + ... + Cn\n• Ketma-ket ulashda zaryadlar teng bo'ladi (Q = Q1 = Q2).\n• Parallel ulashda kuchlanishlar teng bo'ladi (U = U1 = U2).", {
        x: 0.8, y: 2.0, w: 8.4, h: 4.0,
        fontSize: 18, color: 'E2E8F0', lineSpacing: 28, fontFace: 'Calibri'
    });

    pptx.writeFile({ fileName: `${topic}.pptx` });
}
