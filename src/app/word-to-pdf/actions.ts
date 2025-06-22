"use server";

import PDFDocument from 'pdfkit';
import mammoth from 'mammoth';
import type { Buffer } from 'buffer';
import fs from 'fs';
import path from 'path';

export interface ConvertWordToPdfInput {
  docxFileBase64: string;
  originalFileName: string;
}

export interface ConvertWordToPdfOutput {
  pdfDataUri?: string;
  error?: string;
}

// Font configuration for different language support
const FONT_CONFIG = [
  {
    name: 'NotoSansSC',
    file: 'NotoSansSC-VariableFont_wght.ttf',
    description: 'Chinese Simplified',
    unicodeRanges: [
      [0x4E00, 0x9FFF], // CJK Unified Ideographs
      [0x3400, 0x4DBF], // CJK Extension A
      [0x20000, 0x2A6DF], // CJK Extension B
      [0x2A700, 0x2B73F], // CJK Extension C
      [0x2B740, 0x2B81F], // CJK Extension D
      [0x2B820, 0x2CEAF], // CJK Extension E
      [0x2CEB0, 0x2EBEF], // CJK Extension F
      [0x3000, 0x303F], // CJK Symbols and Punctuation
      [0xFF00, 0xFFEF], // Halfwidth and Fullwidth Forms
    ]
  },
  {
    name: 'DejaVuSans',
    file: 'DejaVuSans.ttf',
    description: 'Latin, Cyrillic, Greek',
    unicodeRanges: [
      [0x0000, 0x007F], // Basic Latin
      [0x0080, 0x00FF], // Latin-1 Supplement
      [0x0100, 0x017F], // Latin Extended-A
      [0x0180, 0x024F], // Latin Extended-B
      [0x0370, 0x03FF], // Greek and Coptic
      [0x0400, 0x04FF], // Cyrillic
      [0x0500, 0x052F], // Cyrillic Supplement
      [0x1E00, 0x1EFF], // Latin Extended Additional
      [0x2000, 0x206F], // General Punctuation
      [0x20A0, 0x20CF], // Currency Symbols
    ]
  }
];

class FontManager {
  private loadedFonts: Map<string, Buffer> = new Map();
  private fontPriority: string[] = [];

  constructor() {
    this.loadAvailableFonts();
  }

  private loadAvailableFonts() {
    const fontsDir = path.join(process.cwd(), 'src', 'assets', 'fonts');
    
    for (const fontConfig of FONT_CONFIG) {
      const fontPath = path.join(fontsDir, fontConfig.file);
      try {
        if (fs.existsSync(fontPath)) {
          const fontBuffer = fs.readFileSync(fontPath);
          this.loadedFonts.set(fontConfig.name, fontBuffer);
          this.fontPriority.push(fontConfig.name);
          console.log(`Successfully loaded font: ${fontConfig.name} (${fontConfig.description})`);
        } else {
          console.warn(`Font not found: ${fontPath}`);
        }
      } catch (error) {
        console.error(`Error loading font ${fontConfig.name}:`, error);
      }
    }

    if (this.loadedFonts.size === 0) {
      console.warn('No custom fonts loaded. Will use PDFKit default fonts.');
    }
  }

  public getBestFontForText(text: string): string | Buffer | null {
    if (this.loadedFonts.size === 0) {
      return null; // Use PDFKit default
    }

    // Analyze text to determine required character ranges
    const textCodePoints = Array.from(text).map(char => char.codePointAt(0) || 0);
    
    // Check each font in priority order
    for (const fontName of this.fontPriority) {
      const fontConfig = FONT_CONFIG.find(f => f.name === fontName);
      if (!fontConfig) continue;

      // Check if this font supports the majority of characters in the text
      let supportedChars = 0;
      for (const codePoint of textCodePoints) {
        if (this.isCodePointSupported(codePoint, fontConfig.unicodeRanges)) {
          supportedChars++;
        }
      }

      // If font supports more than 80% of characters, use it
      if (supportedChars / textCodePoints.length > 0.8) {
        return this.loadedFonts.get(fontName) || null;
      }
    }

    // Fallback to first available font
    const firstFont = this.fontPriority[0];
    return firstFont ? this.loadedFonts.get(firstFont) || null : null;
  }

  private isCodePointSupported(codePoint: number, ranges: number[][]): boolean {
    return ranges.some(([start, end]) => codePoint >= start && codePoint <= end);
  }

  public getAvailableFonts(): string[] {
    return Array.from(this.loadedFonts.keys());
  }
}

export async function convertWordToPdfAction(input: ConvertWordToPdfInput): Promise<ConvertWordToPdfOutput> {
  if (!input.docxFileBase64) {
    return { error: "No DOCX file data provided for conversion." };
  }

  console.log(`Converting Word file: ${input.originalFileName} with multi-language font support.`);

  try {
    const docxFileBuffer = Buffer.from(input.docxFileBase64, 'base64');

    if (docxFileBuffer.length === 0) {
      return { error: "Empty DOCX file data received after base64 decoding." };
    }

    // Initialize font manager
    const fontManager = new FontManager();
    console.log(`Available fonts: ${fontManager.getAvailableFonts().join(', ')}`);

    // 1. Extract raw text
    const rawTextResult = await mammoth.extractRawText({ buffer: docxFileBuffer });
    const textContent = rawTextResult.value;

    if (!textContent.trim()) {
      return { error: "No text content found in the document." };
    }

    // 2. Extract images
    const images: { buffer: Buffer; contentType: string }[] = [];
    const imageConvertOptions = {
      convertImage: mammoth.images.imgElement(async (image) => {
        const imageBuffer = await image.read();
        images.push({ buffer: imageBuffer, contentType: image.contentType });
        return {}; // Required return for imgElement
      }),
    };
    await mammoth.convertToHtml({ buffer: docxFileBuffer }, imageConvertOptions);

    // 3. Create PDF with PDFKit
    const pdfDoc = new PDFDocument({ 
      autoFirstPage: false, 
      margins: { top: 72, bottom: 72, left: 72, right: 72 },
      bufferPages: true // Enable page buffering for better performance
    });
    
    const pdfChunks: Buffer[] = [];
    pdfDoc.on('data', (chunk) => pdfChunks.push(chunk as Buffer));
    
    pdfDoc.addPage();

    // 4. Process text with appropriate fonts
    await this.processTextWithFonts(pdfDoc, textContent, fontManager);

    // 5. Add images if any
    if (images.length > 0) {
      await this.processImages(pdfDoc, images, fontManager, input.originalFileName);
    }
    
    return new Promise<ConvertWordToPdfOutput>((resolve, reject) => {
      pdfDoc.on('end', () => {
        const pdfBytes = Buffer.concat(pdfChunks);
        const pdfDataUri = `data:application/pdf;base64,${pdfBytes.toString('base64')}`;
        console.log(`PDF conversion successful for ${input.originalFileName} with multi-language support.`);
        resolve({ pdfDataUri });
      });

      pdfDoc.on('error', (err) => {
        console.error(`Error during PDFKit stream finalization for ${input.originalFileName}:`, err);
        reject({ error: "Failed to finalize PDF document. " + err.message });
      });
      
      pdfDoc.end();
    });

  } catch (e: any) {
    console.error(`Error converting DOCX ${input.originalFileName}:`, e);
    let errorMessage = "Failed to convert Word document. " + e.message;
    if (e.message && e.message.includes("Unrecognised Office Open XML")) {
        errorMessage = "The uploaded file does not appear to be a valid .docx file or is corrupted.";
    }
    return { error: errorMessage };
  }
}

async function processTextWithFonts(pdfDoc: PDFDocument, text: string, fontManager: FontManager) {
  // Split text into paragraphs
  const paragraphs = text.split(/\n\s*\n/).filter(p => p.trim());
  
  for (const paragraph of paragraphs) {
    if (!paragraph.trim()) continue;

    // Get the best font for this paragraph
    const fontBuffer = fontManager.getBestFontForText(paragraph);
    
    try {
      if (fontBuffer) {
        pdfDoc.font(fontBuffer);
      } else {
        // Fallback to PDFKit default
        pdfDoc.font('Helvetica');
      }

      // Add paragraph with proper spacing
      pdfDoc.fontSize(12).text(paragraph.trim(), {
        align: 'left',
        lineGap: 4,
        paragraphGap: 8,
        width: pdfDoc.page.width - 144, // Account for margins
      });

      // Add space between paragraphs
      pdfDoc.moveDown(0.5);

    } catch (fontError) {
      console.warn(`Font error for paragraph, falling back to default:`, fontError);
      pdfDoc.font('Helvetica');
      pdfDoc.fontSize(12).text(paragraph.trim(), {
        align: 'left',
        lineGap: 4,
        paragraphGap: 8,
        width: pdfDoc.page.width - 144,
      });
      pdfDoc.moveDown(0.5);
    }
  }
}

async function processImages(pdfDoc: PDFDocument, images: { buffer: Buffer; contentType: string }[], fontManager: FontManager, fileName: string) {
  for (const img of images) {
    pdfDoc.addPage();
    try {
      if (img.contentType === 'image/jpeg' || img.contentType === 'image/png') {
        pdfDoc.image(img.buffer, {
          fit: [pdfDoc.page.width - 144, pdfDoc.page.height - 144], // Fit within margins
          align: 'center',
          valign: 'center',
        });
      } else {
        console.warn(`Skipping image with unsupported content type: ${img.contentType} in file ${fileName}`);
        
        // Use appropriate font for error message
        const errorText = `[Unsupported image type: ${img.contentType}]`;
        const fontBuffer = fontManager.getBestFontForText(errorText);
        
        if (fontBuffer) {
          pdfDoc.font(fontBuffer);
        } else {
          pdfDoc.font('Helvetica');
        }
        
        pdfDoc.fontSize(10).text(errorText, { align: 'center' });
      }
    } catch (imgError: any) {
      console.error(`Error embedding image in PDFKit for ${fileName}:`, imgError);
      
      // Use appropriate font for error message
      const errorText = `[Error embedding image: ${imgError.message}]`;
      const fontBuffer = fontManager.getBestFontForText(errorText);
      
      if (fontBuffer) {
        pdfDoc.font(fontBuffer);
      } else {
        pdfDoc.font('Helvetica');
      }
      
      pdfDoc.fontSize(10).text(errorText, { align: 'center' });
    }
  }
}