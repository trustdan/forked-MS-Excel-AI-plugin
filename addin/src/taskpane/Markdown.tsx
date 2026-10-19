import React, { useState } from 'react';

interface MarkdownProps {
  content: string;
}

/**
 * Lightweight, zero-dependency, safe Markdown parser & renderer tailored for Excel taskpane.
 * Handles headings, tables, code blocks with copy button, lists, blockquotes, bold, italics,
 * and inline code without using dangerouslySetInnerHTML.
 */
export const Markdown: React.FC<MarkdownProps> = ({ content }) => {
  if (!content) return null;

  const blocks = parseBlocks(content);

  return (
    <div className="md-content">
      {blocks.map((block, idx) => (
        <RenderBlock key={idx} block={block} />
      ))}
    </div>
  );
};

/* ---------------- Block Parsing ---------------- */

type Block =
  | { type: 'heading'; level: number; text: string }
  | { type: 'codeblock'; language: string; code: string }
  | { type: 'table'; headers: string[]; rows: string[][] }
  | { type: 'ul'; items: string[] }
  | { type: 'ol'; items: string[] }
  | { type: 'quote'; text: string }
  | { type: 'paragraph'; text: string };

function parseBlocks(raw: string): Block[] {
  const lines = raw.split(/\r?\n/);
  const blocks: Block[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i]!;

    // 1. Fenced Code Block
    if (line.trim().startsWith('```')) {
      const language = line.trim().slice(3).trim();
      const codeLines: string[] = [];
      i++;
      while (i < lines.length && !lines[i]!.trim().startsWith('```')) {
        codeLines.push(lines[i]!);
        i++;
      }
      i++; // consume closing ```
      blocks.push({ type: 'codeblock', language, code: codeLines.join('\n') });
      continue;
    }

    // 2. Empty Line
    if (!line.trim()) {
      i++;
      continue;
    }

    // 3. Headings (# H1, ## H2, etc.)
    const headingMatch = /^(#{1,6})\s+(.*)$/.exec(line);
    if (headingMatch) {
      blocks.push({
        type: 'heading',
        level: headingMatch[1]!.length,
        text: headingMatch[2]!.trim(),
      });
      i++;
      continue;
    }

    // 4. Blockquote (> quote)
    if (line.trim().startsWith('>')) {
      const quoteLines: string[] = [];
      while (i < lines.length && lines[i]!.trim().startsWith('>')) {
        quoteLines.push(lines[i]!.replace(/^>\s?/, ''));
        i++;
      }
      blocks.push({ type: 'quote', text: quoteLines.join('\n') });
      continue;
    }

    // 5. Table (| a | b |\n|---|---|\n| c | d |)
    if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
      const tableLines: string[] = [];
      while (i < lines.length && lines[i]!.trim().startsWith('|') && lines[i]!.trim().endsWith('|')) {
        tableLines.push(lines[i]!);
        i++;
      }
      if (tableLines.length >= 2 && isTableDivider(tableLines[1]!)) {
        const headers = parseTableRow(tableLines[0]!);
        const rows = tableLines.slice(2).map(parseTableRow);
        blocks.push({ type: 'table', headers, rows });
        continue;
      } else {
        // Fallback: not a valid markdown table
        for (const tl of tableLines) {
          blocks.push({ type: 'paragraph', text: tl });
        }
        continue;
      }
    }

    // 6. Unordered List (- item or * item)
    if (/^\s*[-*]\s+/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^\s*[-*]\s+/.test(lines[i]!)) {
        items.push(lines[i]!.replace(/^\s*[-*]\s+/, ''));
        i++;
      }
      blocks.push({ type: 'ul', items });
      continue;
    }

    // 7. Ordered List (1. item)
    if (/^\s*\d+\.\s+/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^\s*\d+\.\s+/.test(lines[i]!)) {
        items.push(lines[i]!.replace(/^\s*\d+\.\s+/, ''));
        i++;
      }
      blocks.push({ type: 'ol', items });
      continue;
    }

    // 8. Normal Paragraph (accumulate until empty line or special block)
    const pLines: string[] = [];
    while (
      i < lines.length &&
      lines[i]!.trim() &&
      !lines[i]!.trim().startsWith('```') &&
      !lines[i]!.trim().startsWith('#') &&
      !lines[i]!.trim().startsWith('>') &&
      !(lines[i]!.trim().startsWith('|') && lines[i]!.trim().endsWith('|')) &&
      !/^\s*[-*]\s+/.test(lines[i]!) &&
      !/^\s*\d+\.\s+/.test(lines[i]!)
    ) {
      pLines.push(lines[i]!);
      i++;
    }
    blocks.push({ type: 'paragraph', text: pLines.join('\n') });
  }

  return blocks;
}

function isTableDivider(line: string): boolean {
  return /^\|[\s-:]+(\|[\s-:]+)+\|$/.test(line.trim());
}

function parseTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\||\|$/g, '');
  return trimmed.split('|').map((c) => c.trim());
}

/* ---------------- Block Rendering ---------------- */

const RenderBlock: React.FC<{ block: Block }> = ({ block }) => {
  switch (block.type) {
    case 'heading': {
      const HeadingTag = `h${Math.min(block.level + 1, 5)}` as keyof JSX.IntrinsicElements;
      return <HeadingTag className="md-heading"><InlineText text={block.text} /></HeadingTag>;
    }
    case 'codeblock':
      return <CodeBlock code={block.code} language={block.language} />;
    case 'table':
      return (
        <div className="md-table-wrap">
          <table className="md-table">
            <thead>
              <tr>
                {block.headers.map((h, i) => (
                  <th key={i}><InlineText text={h} /></th>
                ))}
              </tr>
            </thead>
            <tbody>
              {block.rows.map((row, rIdx) => (
                <tr key={rIdx}>
                  {row.map((cell, cIdx) => (
                    <td key={cIdx}><InlineText text={cell} /></td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    case 'ul':
      return (
        <ul className="md-list">
          {block.items.map((item, i) => (
            <li key={i}><InlineText text={item} /></li>
          ))}
        </ul>
      );
    case 'ol':
      return (
        <ol className="md-list">
          {block.items.map((item, i) => (
            <li key={i}><InlineText text={item} /></li>
          ))}
        </ol>
      );
    case 'quote':
      return (
        <blockquote className="md-quote">
          <InlineText text={block.text} />
        </blockquote>
      );
    case 'paragraph':
      return (
        <p className="md-p">
          <InlineText text={block.text} />
        </p>
      );
  }
};

/* ---------------- Code Block Component with Copy ---------------- */

const CodeBlock: React.FC<{ code: string; language?: string }> = ({ code, language }) => {
  const [copied, setCopied] = useState(false);

  const copy = (): void => {
    void navigator.clipboard.writeText(code).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  return (
    <div className="md-codeblock">
      <div className="md-codeblock-head">
        <span className="md-codeblock-lang">{language || 'code'}</span>
        <button type="button" className="md-copy-btn" onClick={copy} title="Copy to clipboard">
          {copied ? '✓ Copied' : 'Copy'}
        </button>
      </div>
      <pre>
        <code>{code}</code>
      </pre>
    </div>
  );
};

/* ---------------- Inline Text Formatting ---------------- */

interface TextSegment {
  type: 'text' | 'bold' | 'italic' | 'bold_italic' | 'code' | 'link';
  text: string;
  href?: string;
}

const InlineText: React.FC<{ text: string }> = ({ text }) => {
  const segments = parseInline(text);

  return (
    <>
      {segments.map((seg, i) => {
        switch (seg.type) {
          case 'bold_italic':
            return <strong key={i}><em>{seg.text}</em></strong>;
          case 'bold':
            return <strong key={i}>{seg.text}</strong>;
          case 'italic':
            return <em key={i}>{seg.text}</em>;
          case 'code':
            return <code key={i} className="md-inline-code">{seg.text}</code>;
          case 'link':
            return (
              <span key={i} className="md-link" title={seg.href}>
                {seg.text}
              </span>
            );
          default:
            return <React.Fragment key={i}>{seg.text}</React.Fragment>;
        }
      })}
    </>
  );
};

function parseInline(raw: string): TextSegment[] {
  const segments: TextSegment[] = [];
  // Regex to match inline patterns:
  // 1. Inline code: `code`
  // 2. Bold italic: ***text***
  // 3. Bold: **text** or __text__
  // 4. Italic: *text* or _text_
  // 5. Link: [text](href)
  const regex = /(`([^`]+)`)|(\*\*\*([^*]+)\*\*\*)|(\*\*([^*]+)\*\*)|(\*([^*]+)\*)|(__([^_]+)__)|(_([^_]+)_)|(\[([^\]]+)\]\(([^)]+)\))/g;

  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = regex.exec(raw)) !== null) {
    if (match.index > lastIndex) {
      segments.push({ type: 'text', text: raw.slice(lastIndex, match.index) });
    }

    if (match[1]) {
      // `code`
      segments.push({ type: 'code', text: match[2]! });
    } else if (match[3]) {
      // ***bold italic***
      segments.push({ type: 'bold_italic', text: match[4]! });
    } else if (match[5]) {
      // **bold**
      segments.push({ type: 'bold', text: match[6]! });
    } else if (match[7]) {
      // *italic*
      segments.push({ type: 'italic', text: match[8]! });
    } else if (match[9]) {
      // __bold__
      segments.push({ type: 'bold', text: match[10]! });
    } else if (match[11]) {
      // _italic_
      segments.push({ type: 'italic', text: match[12]! });
    } else if (match[13]) {
      // [text](href)
      segments.push({ type: 'link', text: match[14]!, href: match[15]! });
    }

    lastIndex = regex.lastIndex;
  }

  if (lastIndex < raw.length) {
    segments.push({ type: 'text', text: raw.slice(lastIndex) });
  }

  return segments;
}
