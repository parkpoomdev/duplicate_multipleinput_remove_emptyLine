/**
 * สร้างเมนูบน Google Slides (ไม่ต้อง Run เอง)
 */
function onOpen(e) {
  try {
    const ui = SlidesApp.getUi();
    ui.createMenu('Custom Tools')
      .addItem('1) รวม Text เป็น "-" ในสไลด์นี้', 'buildBulletsOnCurrentSlide')
      .addItem('2) รวม Text เป็น "-" แล้ว Duplicate สไลด์นี้', 'mergeAndDuplicateCurrentSlide')
      .addItem('3) Duplicate สไลด์นี้ (รวม Text แบบย่อหน้า, ไม่มี "-")', 'duplicateSlideWithParagraphs')
      .addItem('4) ลบบรรทัดว่างในข้อความที่เลือก', 'removeEmptyLinesInSelection')
      .addToUi();
  } catch (err) {
    Logger.log('UI not available in this context. ' + err);
  }
}

/**
 * เอา "สไลด์ที่เลือกอยู่" (ต้องเป็น SLIDE ปกติ)
 */
function getCurrentSlide_() {
  const pres = SlidesApp.getActivePresentation();
  let selection;
  try {
    selection = pres.getSelection();
  } catch (err) {
    throw new Error('ต้องเรียกใช้จากใน Google Slides เท่านั้น');
  }

  const page = selection.getCurrentPage();
  if (!page) {
    throw new Error('โปรดเลือกสไลด์ก่อนใช้งานเมนู');
  }

  if (page.getPageType() !== SlidesApp.PageType.SLIDE) {
    throw new Error('โปรดเลือกสไลด์ปกติ (ไม่ใช่ Master / Layout / Notes)');
  }

  return page.asSlide();
}

/**
 * หา Title ของสไลด์
 * 1) ถ้ามี placeholder เป็น TITLE / CENTERED_TITLE ใช้อันนั้น
 * 2) ถ้าไม่มี ใช้ text box ที่อยู่ "บนสุด" (top น้อยสุด)
 */
function getTitleShape_(slide) {
  const elements = slide.getPageElements();
  let titleShape = null;

  // ลองหา placeholder title ก่อน
  elements.forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

    const shape = pe.asShape();
    if (shape.getShapeType() !== SlidesApp.ShapeType.TEXT_BOX) return;

    try {
      const placeholder = shape.getPlaceholder();
      if (!placeholder) return;

      const type = placeholder.getType();
      if (type === SlidesApp.PlaceholderType.TITLE ||
          type === SlidesApp.PlaceholderType.CENTERED_TITLE) {
        titleShape = shape;
      }
    } catch (e) {
      // shape ไม่มี placeholder ก็ไม่เป็นไร ข้ามไป
    }
  });

  if (titleShape) return titleShape;

  // ถ้าไม่มี placeholder title เลย: ใช้ text box ที่ top น้อยสุด
  let minTop = null;
  elements.forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

    const shape = pe.asShape();
    if (shape.getShapeType() !== SlidesApp.ShapeType.TEXT_BOX) return;

    const top = shape.getTop();
    if (minTop === null || top < minTop) {
      minTop = top;
      titleShape = shape;
    }
  });

  return titleShape;
}

/**
 * คืน list ของ shape ที่เป็น input text:
 * - เป็น text box
 * - ไม่ใช่ Title
 * - ไม่ว่าง
 * - ไม่ขึ้นต้นด้วย "- "
 */
function getInputTextShapes_(slide) {
  const elements = slide.getPageElements();
  const inputs = [];
  const titleShape = getTitleShape_(slide);
  const titleId = titleShape ? titleShape.getObjectId() : null;

  elements.forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

    const shape = pe.asShape();
    if (shape.getShapeType() !== SlidesApp.ShapeType.TEXT_BOX) return;

    // ข้าม title
    if (titleId && shape.getObjectId() === titleId) return;

    const textRange = shape.getText();
    const text = textRange.asString().trim();
    if (!text) return;
    if (text.startsWith('- ')) return; // ข้ามกล่อง bullet เดิม

    inputs.push(shape);
  });

  return inputs;
}

/**
 * สำหรับจุดที่ต้องการแค่ข้อความอย่างเดียว
 */
function getInputTexts_(slide) {
  return getInputTextShapes_(slide).map(s => s.getText().asString().trim());
}

/**
 * หา text box แรกที่ขึ้นต้นด้วย "- " (เอาไว้ใช้เป็น bullet box เดิม)
 */
function findBulletShape_(slide) {
  const elements = slide.getPageElements();
  let bulletShape = null;

  elements.forEach(pe => {
    if (pe.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = pe.asShape();
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        const text = shape.getText().asString().trim();
        if (text.startsWith('- ') && !bulletShape) {
          bulletShape = shape;
        }
      }
    }
  });

  return bulletShape;
}

/**
 * ลบ text box ทั้งหมดบน slide
 * แต่:
 *  - ไม่ยุ่งกับ Title
 *  - ถ้ามี keepObjectIds ให้กันไว้ไม่ลบ
 */
function clearTextBoxesOnSlide_(slide, options) {
  options = options || {};
  const keepObjectIds = options.keepObjectIds || [];

  const keepSet = {};
  keepObjectIds.forEach(id => {
    if (id) keepSet[id] = true;
  });

  const titleShape = getTitleShape_(slide);
  if (titleShape) {
    keepSet[titleShape.getObjectId()] = true;
  }

  const elements = slide.getPageElements();
  const toRemove = [];

  elements.forEach(pe => {
    if (pe.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = pe.asShape();
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        const id = shape.getObjectId();
        if (!keepSet[id]) {
          toRemove.push(shape);
        }
      }
    }
  });

  toRemove.forEach(shape => shape.remove());
}

/**
 * เมนู 1:
 * รวม text ในสไลด์ปัจจุบัน → เขียน bullet "-" บนสไลด์เดิม (ไม่ duplicate)
 * - input: text box ทุกอันที่ไม่ใช่ Title และไม่ขึ้นต้นด้วย "- "
 * - ถ้ามี bullet box เดิม (เริ่มด้วย "- ") จะเขียนทับในกล่องนั้น
 * - ถ้าไม่มี bullet box เดิม ใช้กล่อง input อันแรกเป็นกล่องรวม
 */
function buildBulletsOnCurrentSlide() {
  const slide = getCurrentSlide_();
  const inputShapes = getInputTextShapes_(slide);
  const inputs = inputShapes.map(s => s.getText().asString().trim());

  if (inputs.length === 0) {
    SlidesApp.getUi().alert('ไม่พบข้อความ input (text box ที่ไม่ขึ้นต้นด้วย "- ") บนสไลด์นี้');
    return;
  }

  const bulletText = inputs.map(t => '- ' + t).join('\n\n');
  let bulletShape = findBulletShape_(slide);

  // ถ้าไม่มี bullet box เดิม ใช้กล่อง input อันแรกเป็นที่เขียนทับ
  if (!bulletShape) {
    bulletShape = inputShapes[0];
  }

  bulletShape.getText().setText(bulletText);
}

/**
 * เมนู 2:
 * รวม text เป็น bullet "-" จาก "สไลด์ต้นฉบับ"
 * → Duplicate สไลด์
 * → บนสไลด์ใหม่:
 *    - เก็บ Title ไว้เหมือนเดิม
 *    - เก็บ text box ตัวแรกของ input ไว้เป็นตู้รวมข้อความ
 *    - ลบ text box อื่นในสไลด์ใหม่
 *    - เขียน bullet text ลงใน text box นั้น (เก็บฟอนต์/สไตล์จากกล่องเดิม)
 */
function mergeAndDuplicateCurrentSlide() {
  const baseSlide = getCurrentSlide_();
  const baseInputShapes = getInputTextShapes_(baseSlide);
  const inputs = baseInputShapes.map(s => s.getText().asString().trim());

  if (inputs.length === 0) {
    SlidesApp.getUi().alert('ไม่พบข้อความ input บนสไลด์นี้');
    return;
  }

  const bulletText = inputs.map(t => '- ' + t).join('\n\n');

  const newSlide = baseSlide.duplicate();

  // หารายการ input shapes บนสไลด์ใหม่
  const newInputShapes = getInputTextShapes_(newSlide);
  if (newInputShapes.length === 0) {
    SlidesApp.getUi().alert('สไลด์ใหม่ไม่พบ input text box (ผิดปกติ)');
    return;
  }

  const targetShape = newInputShapes[0];

  // ลบ text box อื่นทั้งหมด ยกเว้น title + targetShape
  clearTextBoxesOnSlide_(newSlide, {
    keepObjectIds: [targetShape.getObjectId()]
  });

  // ใส่ข้อความ bullet ลงไป ใช้ฟอร์แมตจาก targetShape เดิม
  targetShape.getText().setText(bulletText);
}

/**
 * เมนู 3:
 * Duplicate สไลด์ → บนสไลด์ใหม่:
 *  - เก็บ Title เดิม
 *  - เก็บ text box input อันแรกไว้เป็นกล่องหลัก
 *  - ลบ text box input อื่นออก
 *  - รวม input ทุกอัน เป็นย่อหน้า คั่นด้วยบรรทัดว่าง (ไม่มี "-")
 */
function duplicateSlideWithParagraphs() {
  const baseSlide = getCurrentSlide_();
  const baseInputShapes = getInputTextShapes_(baseSlide);
  const inputs = baseInputShapes.map(s => s.getText().asString().trim());

  if (inputs.length === 0) {
    SlidesApp.getUi().alert('ไม่พบข้อความ input บนสไลด์นี้');
    return;
  }

  const paragraphText = inputs.join('\n\n');

  const newSlide = baseSlide.duplicate();
  const newInputShapes = getInputTextShapes_(newSlide);
  if (newInputShapes.length === 0) {
    SlidesApp.getUi().alert('สไลด์ใหม่ไม่พบ input text box (ผิดปกติ)');
    return;
  }

  const targetShape = newInputShapes[0];

  clearTextBoxesOnSlide_(newSlide, {
    keepObjectIds: [targetShape.getObjectId()]
  });

  targetShape.getText().setText(paragraphText);
}

/**
 * เมนู 4:
 * ลบบรรทัดว่างใน "ข้อความที่เลือกอยู่"
 * - ใช้ selection.getTextRange()
 * - เอาเฉพาะบรรทัดที่มีตัวอักษร (trim แล้วไม่ว่าง)
 * - ต่อกลับด้วย '\n'
 */
function removeEmptyLinesInSelection() {
  const pres = SlidesApp.getActivePresentation();
  let selection;
  try {
    selection = pres.getSelection();
  } catch (err) {
    throw new Error('ต้องเรียกใช้จากใน Google Slides เท่านั้น');
  }

  const textRange = selection.getTextRange();
  if (!textRange) {
    SlidesApp.getUi().alert('โปรดเลือกข้อความก่อน (ลากเลือกช่วงที่ต้องการลบบรรทัดว่าง)');
    return;
  }

  const original = textRange.asString();
  const lines = original.split(/\r?\n/);

  const filtered = lines.filter(line => line.trim() !== '');
  const newText = filtered.join('\n');

  textRange.setText(newText);
}
