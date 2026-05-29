const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const pptPath = 'C:\\Users\\KITC\\Desktop\\바이오파운드리 ISMP\\0. ISMP 산출물\\01. 사업의 개요\\바이오파운드리 ISMP_Ⅰ. 사업 개요_v0.6.pptx';

async function run() {
    if (!fs.existsSync(pptPath)) {
        console.error("PPT 파일이 존재하지 않습니다:", pptPath);
        return;
    }
    const data = fs.readFileSync(pptPath);
    const zip = await JSZip.loadAsync(data);
    
    const sizes = new Set();
    const slides = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
    
    console.log(`총 ${slides.length}개의 슬라이드를 분석합니다...`);
    
    for (const slide of slides) {
        const text = await zip.files[slide].async('text');
        
        // 1. sz="value" 속성 수집
        const matches = text.match(/sz="(\d+)"/g);
        if (matches) {
            matches.forEach(m => {
                const val = m.match(/\d+/)[0];
                sizes.add(parseInt(val));
            });
        }
    }
    
    const sortedSizes = Array.from(sizes).sort((a,b)=>a-b);
    console.log("PPT 내부에서 발견된 모든 sz 속성값들 (sz / 100 pt):");
    sortedSizes.forEach(sz => {
        console.log(`sz=${sz} (${sz/100}pt)`);
    });
}

run().catch(console.error);
