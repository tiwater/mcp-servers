import assert from 'node:assert/strict';
import test from 'node:test';
import { compactDocxObjectIdentity } from './docx-object-identity.mjs';

test('normalizes omitted nullable DOCX identity fields', () => {
  assert.deepEqual(compactDocxObjectIdentity({
    address: { part: '/word/header1.xml', path: '/w:hdr[1]/w:p[1]' },
    parentAddress: { part: '/word/header1.xml', path: '/w:hdr[1]' },
    kind: 'paragraph',
    textPreview: 'Header',
  }), {
    address: { part: '/word/header1.xml', path: '/w:hdr[1]/w:p[1]' },
    parentAddress: { part: '/word/header1.xml', path: '/w:hdr[1]' },
    kind: 'paragraph',
    textPreview: 'Header',
    gridSpan: null,
    verticalMerge: null,
    verticalTextAlignment: null,
  });
});

test('preserves native DOCX identity values', () => {
  const address = { part: '/word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]' };
  assert.deepEqual(compactDocxObjectIdentity({
    address,
    parentAddress: null,
    kind: 'run',
    textPreview: null,
    gridSpan: 2,
    verticalMerge: 'restart',
    verticalTextAlignment: 'superscript',
  }), {
    address,
    parentAddress: null,
    kind: 'run',
    textPreview: null,
    gridSpan: 2,
    verticalMerge: 'restart',
    verticalTextAlignment: 'superscript',
  });
});
