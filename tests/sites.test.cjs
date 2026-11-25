// node --test tests/sites.test.cjs : gs_sitesOfUser (js/sharepoint_graph_common.js) with Graph answered by fixtures.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const src = fs.readFileSync(path.join(__dirname, '..', 'js', 'sharepoint_graph_common.js'), 'utf8');
const start = src.indexOf('function gs_sitesOfUser(');
const end = src.indexOf('\nfunction gs_graphSiteListColumnsGetRequestInfo(');
const BASE = 'https://graph.microsoft.com/v1.0';
const site = (id, name, url) => ({ id, displayName: name, name, webUrl: url });
function load(answers, failSearch) {
  const calls = [];
  const env = {
    gs_safeString: (v) => (v === null || v === undefined ? '' : String(v)),
    gs_getGraphClientContext: () => ({ adapter: { getBaseUrl: () => BASE } }),
    gs_normalizeSite: (s) => ({ siteId: s.id, id: s.id, displayName: s.displayName, name: s.name, webUrl: s.webUrl }),
    gs_graphSitesSearchUrl: (t, q) => BASE + '/sites?search=' + q,
    gs_httpStatusOf: (e) => ({ status: e.httpStatus === undefined ? null : e.httpStatus }),
    gs_graphGetJson: (url) => { calls.push(url); if (!(url in answers)) { const e = new Error('404'); e.httpStatus = 404; throw e; } return answers[url]; },
    gs_graphCollectPagedValues: (url) => {
      calls.push(url);
      if (url.indexOf('/sites?search=') !== -1 && failSearch) { const e = new Error('500'); e.httpStatus = 500; throw e; }
      if (!(url in answers)) { const e = new Error('404'); e.httpStatus = 404; throw e; }
      return { values: answers[url].value };
    }
  };
  const fn = new Function(...Object.keys(env), src.substring(start, end) + '\nreturn gs_sitesOfUser;')(...Object.values(env));
  return { fn, calls };
}
const answers = {
  [BASE + '/me/followedSites?$select=id,displayName,webUrl']: { value: [site('s1', 'Test import', 'https://c.sharepoint.com/sites/Test'), site('s2', 'Ventes', 'https://c.sharepoint.com/sites/Ventes')] },
  [BASE + '/sites/root?$select=id,displayName,webUrl']: site('root', 'Communication', 'https://c.sharepoint.com'),
  [BASE + '/sites?search=*']: { value: [site('s1', 'Test import', 'https://c.sharepoint.com/sites/Test'), site('s3', 'RH', 'https://c.sharepoint.com/sites/RH')] },
  [BASE + '/sites?search=vent']: { value: [] }
};
const fields = ['id', 'displayName', 'webUrl'];

test('the sites a user opens without searching: followed, the root site, then the search, each once', () => {
  const { fn } = load(answers, false);
  const found = JSON.parse(JSON.stringify(fn('t', '', fields, 1)));
  assert.deepEqual(found.items.map((s) => s.id), ['s1', 's2', 'root', 's3']);
  assert.equal(found.followed, 2);
  assert.equal(found.searchFailed, '');
});

test('the search of Graph failing (500 on some tenants): the followed and root sites still, filtered by the text', () => {
  const { fn } = load(answers, true);
  let found = JSON.parse(JSON.stringify(fn('t', '*', fields, 1)));
  assert.deepEqual(found.items.map((s) => s.id), ['s1', 's2', 'root']);
  assert.equal(found.searchFailed, '500');
  found = JSON.parse(JSON.stringify(fn('t', 'VENT', fields, 1)));
  assert.deepEqual(found.items.map((s) => s.id), ['s2']);
});

test('an application token (no user): no followed sites, the others still', () => {
  const { fn } = load({ [BASE + '/sites/root?$select=id,displayName,webUrl']: answers[BASE + '/sites/root?$select=id,displayName,webUrl'], [BASE + '/sites?search=*']: answers[BASE + '/sites?search=*'] }, false);
  const found = JSON.parse(JSON.stringify(fn('t', '', fields, 1)));
  assert.deepEqual(found.items.map((s) => s.id), ['root', 's1', 's3']);
  assert.equal(found.followed, 0);
});
