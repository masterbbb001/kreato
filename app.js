function generatePPTX() {
    let topic = document.getElementById('topic').value || "Mavzu ko'rsatilmadi";
    let author = document.getElementById('author').value || "Muallif";

    let pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    // 1-SLAYD: Titul (Ultra Premium Fon)
    let slide1 = pptx.addSlide();
    slide1.background = { color: "0F172A" };
    
    slide1.addText(topic.toUpperCase(), { 
        x: 0.8, y: 2.2, w: 8.5, h: 1.5, 
        fontSize: 36, bold: true, color: "38BDF8", align: "left" 
    });
    
    slide1.addText("Tayyorladi: " + author, { 
        x: 0.8, y: 4.0, w: 8.5, h: 0.8, 
        fontSize: 20, color: "94A3B8", align: "left" 
    });

    // 2-SLAYD: Reja / Mundarija
    let slide2 = pptx.addSlide();
    slide2.background = { color: "1E293B" };
    
    slide2.addText("MAVZU REJASI", { 
        x: 0.8, y: 0.8, w: 8.5, h: 0.8, 
        fontSize: 28, bold: true, color: "F8FAFC" 
    });

    slide2.addShape(pptx.Shapes.rect, { x: 0.8, y: 2.0, w: 8.4, h: 1.2, fill: { color: "334155" } });
    slide2.addText("1. Kirish va asosiy tushunchalar", { x: 1.0, y: 2.2, fontSize: 18, color: "38BDF8" });

    slide2.addShape(pptx.Shapes.rect, { x: 0.8, y: 3.5, w: 8.4, h: 1.2, fill: { color: "334155" } });
    slide2.addText("2. Amaliy tahlil va asosiy faktlar", { x: 1.0, y: 3.7, fontSize: 18, color: "38BDF8" });

    // Faylni yuklab olish
    pptx.writeFile({ fileName: topic + "_Kreato.pptx" });
}
