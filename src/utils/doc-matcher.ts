/**
 * Uses OpenAI to intelligently match document files to GDRFA upload slot labels.
 *
 * Given a list of file names from the applicant's documents folder and the
 * available upload slot labels on the portal, the LLM determines the best match.
 */

import OpenAI from 'openai';

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';

export interface DocSlotMatch {
  slotLabel: string;   // The upload slot on the portal (e.g. "Sponsored Passport page 1")
  fileName: string;    // The matched file name (e.g. "Sponsored Passport page 1.jpg")
  confidence: string;  // "high", "medium", "low"
}

/**
 * Uses OpenAI GPT to match document files to upload slots.
 *
 * @param fileNames - Array of file names in the applicant's documents folder
 * @param slotLabels - Array of available upload slot labels on the portal
 * @param applicantName - Name of the applicant (for context)
 * @returns Array of matched slot→file pairs
 */
export async function matchDocumentsToSlots(
  fileNames: string[],
  slotLabels: string[],
  applicantName: string
): Promise<DocSlotMatch[]> {
  if (!OPENAI_API_KEY) {
    console.warn('[DocMatcher] OPENAI_API_KEY not set — skipping AI matching.');
    return [];
  }

  const openai = new OpenAI({ apiKey: OPENAI_API_KEY });

  const prompt = `You are a document matching assistant for a UAE visa application portal (GDRFA).

Given the applicant "${applicantName}", match each upload slot to the best matching file from the documents folder.

**Upload slots available on the portal:**
${slotLabels.map((s, i) => `${i + 1}. ${s}`).join('\n')}

**Files in the documents folder:**
${fileNames.map((f, i) => `${i + 1}. ${f}`).join('\n')}

**Matching rules:**
- "Sponsored Passport page 1" → match files containing "sponsored passport" or "passport page 1" or the main passport scan
- "Passport External Cover Page" → match files with "external cover", "cover page", or "passport cover"
- "Personal Photo" → match files with "personal photo", "photo", "selfie"
- "Hotel reservation/Place of stay - Page 1" → match files with "hotel", "reservation", "tenancy contract", "accommodation"
- "Return air ticket - Page 1" → match files with "flight ticket", "air ticket", "ticket", "flight", "boarding pass"
- "Others Page 1" → use for hotel reservation if no dedicated hotel slot exists
- "Others Page 2" → use for flight ticket if no dedicated ticket slot exists
- "Birth Certificate Page 1" → match files with "birth certificate"
- "Sponsored Passport page 2-5" → match additional passport page scans
- Parent passport/visa pages → match files with "father passport", "mother passport", "parent"

Respond ONLY with a valid JSON array. Each element must have:
- "slotLabel": exact upload slot label from the list above
- "fileName": exact file name from the list above
- "confidence": "high", "medium", or "low"

Only include matches where you are reasonably confident. Skip slots with no matching file.
Do not invent file names. Use exact file names from the provided list.`;

  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.1,
      max_tokens: 2000,
    });

    const content = response.choices[0]?.message?.content?.trim() || '[]';
    // Extract JSON from the response (might be wrapped in ```json...```)
    const jsonStr = content.replace(/^```json?\s*/, '').replace(/\s*```$/, '').trim();
    const matches: DocSlotMatch[] = JSON.parse(jsonStr);

    console.log(`[DocMatcher] OpenAI matched ${matches.length} documents for "${applicantName}".`);
    for (const m of matches) {
      console.log(`  ${m.slotLabel} → ${m.fileName} (${m.confidence})`);
    }

    return matches;
  } catch (error) {
    console.error('[DocMatcher] OpenAI API error:', error);
    return [];
  }
}

/**
 * Fallback: rule-based matching without LLM (used if OpenAI fails).
 */
export function matchDocumentsToSlotsLocal(
  fileNames: string[],
  slotLabels: string[]
): DocSlotMatch[] {
  const matches: DocSlotMatch[] = [];
  const usedFiles = new Set<string>();

  const rules: Array<{ slotKeywords: string[]; fileKeywords: string[] }> = [
    { slotKeywords: ['sponsored passport page 1'], fileKeywords: ['sponsored passport page 1', 'sponsored passport'] },
    { slotKeywords: ['sponsored passport page 2'], fileKeywords: ['sponsored passport page 2'] },
    { slotKeywords: ['passport external cover'], fileKeywords: ['external cover', 'cover page', 'passport cover'] },
    { slotKeywords: ['personal photo'], fileKeywords: ['personal photo', 'photo'] },
    { slotKeywords: ['hotel reservation', 'place of stay'], fileKeywords: ['hotel', 'reservation', 'tenancy contract', 'tenancy', 'accommodation'] },
    { slotKeywords: ['return air ticket'], fileKeywords: ['flight ticket', 'air ticket', 'ticket', 'flight', 'boarding'] },
    { slotKeywords: ['birth certificate'], fileKeywords: ['birth certificate'] },
    { slotKeywords: ['father passport'], fileKeywords: ['father passport'] },
    { slotKeywords: ['mother passport'], fileKeywords: ['mother passport'] },
  ];

  for (const slot of slotLabels) {
    const slotLower = slot.toLowerCase();
    for (const rule of rules) {
      if (!rule.slotKeywords.some(kw => slotLower.includes(kw))) continue;
      for (const fn of fileNames) {
        if (usedFiles.has(fn)) continue;
        const fnLower = fn.toLowerCase();
        if (rule.fileKeywords.some(kw => fnLower.includes(kw))) {
          matches.push({ slotLabel: slot, fileName: fn, confidence: 'high' });
          usedFiles.add(fn);
          break;
        }
      }
      break;
    }
  }

  return matches;
}
