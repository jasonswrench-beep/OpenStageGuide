import test from 'node:test';
import assert from 'node:assert/strict';
import {MODELS,SOURCES,REFERENCE_SOURCES,modelById,sourceFor,referenceText,inTextCitation} from '../src/design-models.mjs';

test('every model has a name, who it is by, what it is for, events with ideas, and a source that exists', () => {
  assert.ok(MODELS.length>=9);
  assert.equal(new Set(MODELS.map(m=>m.id)).size,MODELS.length);
  for(const m of MODELS){
    assert.ok(m.name&&m.by&&m.kind&&m.summary.length>40&&m.useWhen.length>20,m.id);
    assert.ok(m.events.length>=3,m.id);
    for(const e of m.events){assert.ok(e.name);assert.ok(e.ideas.length>=1&&e.ideas.every(i=>i.length>5),`${m.id}: ${e.name}`);}
    assert.ok(SOURCES[m.source],`${m.id} names a source that is there`);
  }
  assert.equal(modelById('gagne').events.length,9,'nine events');
  assert.equal(modelById('merrill').events.length,5);
  assert.equal(modelById('arcs').events.length,4);
  assert.equal(modelById('nope'),null);
  assert.equal(sourceFor('gagne').year,1985);
  assert.equal(sourceFor('mayer2009').year,2009,'a source by its own id');
  assert.equal(sourceFor('zzz'),null);
});

test('each source has what an APA 7 entry needs, and the DOIs are the ones that were confirmed', () => {
  for(const [id,s] of Object.entries(SOURCES)){
    assert.ok(s.authors.length&&s.authors.every(a=>a.family&&/^[A-Z]/.test(a.given)),id);
    assert.ok(Number.isInteger(s.year)&&s.year>1950&&s.year<2027,id);
    assert.ok(s.title&&(s.type==='article-journal'?s.container&&s.volume&&s.pages:s.publisher),id);
    if(s.doi)assert.match(s.doi,/^10\.\d{4,9}\/\S+$/,id);
  }
  const withDoi=Object.entries(SOURCES).filter(([,s])=>s.doi).map(([id])=>id).sort();
  assert.deepEqual(withDoi,['branch2009','keller1987','krathwohl2002','mayer2009','merrill2002','sweller1988','vanmerrienboer2018']);
  for(const r of REFERENCE_SOURCES)for(const id of r.sources)assert.ok(SOURCES[id],id);
});

test('the reference-list entry follows APA 7 for an article, a book with an edition, an edited book, a book with two authors and a report', () => {
  assert.equal(referenceText(SOURCES.merrill2002),'Merrill, M. D. (2002). First principles of instruction. Educational Technology Research and Development, 50(3), 43–59. https://doi.org/10.1007/BF02505024');
  assert.equal(referenceText(SOURCES.gagne1985),'Gagné, R. M. (1985). The conditions of learning and theory of instruction (4th ed.). Holt, Rinehart and Winston.');
  assert.equal(referenceText(SOURCES.anderson2001),'Anderson, L. W., & Krathwohl, D. R. (Eds.). (2001). A taxonomy for learning, teaching, and assessing: A revision of Bloom\'s taxonomy of educational objectives. Longman.');
  assert.equal(referenceText(SOURCES.wiggins2005),'Wiggins, G., & McTighe, J. (2005). Understanding by design (Expanded 2nd ed.). Association for Supervision and Curriculum Development.');
  assert.equal(referenceText(SOURCES.bybee2006),'Bybee, R. W., Taylor, J. A., Gardner, A., Van Scotter, P., Powell, J. C., Westbrook, A., & Landes, N. (2006). The BSCS 5E instructional model: Origins and effectiveness. BSCS.');
  assert.equal(referenceText(SOURCES.kolb1984),'Kolb, D. A. (1984). Experiential learning: Experience as the source of learning and development. Prentice-Hall.');
});

test('the citation in a sentence names one author, two with an ampersand, or the first and "et al."', () => {
  assert.equal(inTextCitation(SOURCES.merrill2002),'(Merrill, 2002)');
  assert.equal(inTextCitation(SOURCES.wiggins2005),'(Wiggins & McTighe, 2005)');
  assert.equal(inTextCitation(SOURCES.bybee2006),'(Bybee et al., 2006)');
});
