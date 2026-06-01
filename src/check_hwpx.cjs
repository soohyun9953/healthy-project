const hwpx = require('hwpx-js');

async function test() {
    const builder = new hwpx.HWPXBuilder();
    builder.addParagraph("테스트 제목", { fontSize: 1600, bold: true });
    builder.addParagraph("테스트 본문 내용입니다.");
    
    const doc = builder.build();
    
    // write 가동
    const result = await hwpx.write(doc);
    console.log("Write Result Type:", result ? result.constructor.name : "null");
    console.log("Is Buffer?:", Buffer.isBuffer(result));
    console.log("Is Uint8Array?:", result instanceof Uint8Array);
    console.log("Result length:", result ? (result.length || result.byteLength) : "none");
}

test().catch(console.error);
