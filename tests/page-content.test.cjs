// node --test tests/ : js/sharepoint_page_content.js with Graph answered by fixtures (no Java, no network).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SITE = 'contoso.sharepoint.com,11111111-1111-4111-8111-111111111111,22222222-2222-4222-8222-222222222222';
const BASE = 'https://graph.microsoft.com/v1.0';
const fixture = (name) => JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', name), 'utf8'));

// the library loaded with the helpers of sharepoint_graph_common.js it uses, Graph answered by a map of URLs
function load(answers, downloads) {
  const calls = [];
  const context = {
    calls,
    gs_safeString: (v) => (v === null || v === undefined ? '' : String(v)),
    gs_isBlank: (v) => v === null || v === undefined || String(v).trim().length === 0,
    gs_require: (n, v) => { if (v === null || v === undefined || String(v).trim() === '') { throw new Error(n + ' is required'); } return String(v); },
    gs_encodePathSegment: (v) => encodeURIComponent(String(v)),
    gs_getGraphClientContext: () => ({ adapter: { getBaseUrl: () => BASE } }),
    gs_httpFailure: (message, r) => { const e = new Error(message); e.httpStatus = r.statusCode; return e; },
    gs_httpStatusOf: (e) => ({ status: e && e.httpStatus !== undefined ? e.httpStatus : null }),
    gs_restSiteUrl: (v) => { if (!/^https:\/\/[^/]+\/./.test(String(v))) { throw new Error('Invalid siteUrl'); } return String(v).replace(/\/+$/, ''); },
    gs_graphGetJson: (url) => {
      calls.push(url);
      if (!(url in answers)) { const e = new Error('404 ' + url); e.httpStatus = 404; throw e; }
      return JSON.parse(JSON.stringify(answers[url]));
    },
    gs_graphCollectPagedValues: (url) => {
      calls.push(url);
      if (!(url in answers)) { const e = new Error('404 ' + url); e.httpStatus = 404; throw e; }
      return { values: JSON.parse(JSON.stringify(answers[url].value)) };
    }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'js', 'sharepoint_page_content.js'), 'utf8'), context);
  context.gs_pageDownloadBytes = (url, token, max) => {
    calls.push(url);
    const d = downloads[url];
    if (!d) { const e = new Error('404'); e.httpStatus = 404; throw e; }
    if (d.length > max) { const e = new Error('413'); e.httpStatus = 413; throw e; }
    return { bytes: d, contentType: 'image/png' };
  };
  context.gs_pageBase64 = (bytes) => Buffer.from(bytes).toString('base64');
  return context;
}

const PNG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3];
const SITE_URL = BASE + '/sites/' + encodeURIComponent(SITE);
const NEWS_URL = SITE_URL + '/pages/microsoft.graph.sitePage?$select=id,title,description,webUrl,thumbnailWebUrl,promotionKind,createdDateTime&$top=100';
const answersOf = (page) => ({
  [SITE_URL + '?$select=id,webUrl,displayName']: { id: SITE, webUrl: 'https://contoso.sharepoint.com/sites/Team', displayName: 'Team' },
  [SITE_URL + '/lists?$select=id,name,displayName,webUrl,list,system&$top=200']: { value: [
    { id: 'aaaaaaaa-0000-4000-8000-000000000001', name: 'Clients', displayName: 'Clients', webUrl: 'https://contoso.sharepoint.com/sites/Team/Lists/Clients', list: { template: 'genericList' } },
    { id: 'aaaaaaaa-0000-4000-8000-000000000002', name: 'Shared Documents', displayName: 'Documents', webUrl: 'https://contoso.sharepoint.com/sites/Team/Shared%20Documents', list: { template: 'documentLibrary' } },
    { id: 'aaaaaaaa-0000-4000-8000-000000000003', name: 'SiteAssets', displayName: 'Site Assets', webUrl: 'https://contoso.sharepoint.com/sites/Team/SiteAssets', list: { template: 'documentLibrary' } }
  ] },
  [SITE_URL + '/pages/microsoft.graph.sitePage?$select=id,name,title,pageLayout,webUrl']: { value: [
    { id: 'p-article', name: 'News.aspx', title: 'News', pageLayout: 'article' },
    { id: 'p-home', name: 'Home.aspx', title: 'Home', pageLayout: 'home', webUrl: 'https://contoso.sharepoint.com/sites/Team/SitePages/Home.aspx' }
  ] },
  [SITE_URL + '/pages/p-home/microsoft.graph.sitePage?$expand=canvasLayout']: page,
  [SITE_URL + '/drive?$select=id,webUrl']: { id: 'drive-docs', webUrl: 'https://contoso.sharepoint.com/sites/Team/Shared%20Documents' },
  // Graph leaves Site Assets out of /drives: its drive by its list
  [SITE_URL + '/drives?$select=id,webUrl']: { value: [
    { id: 'drive-docs', webUrl: 'https://contoso.sharepoint.com/sites/Team/Shared%20Documents' }
  ] },
  [SITE_URL + '/lists/aaaaaaaa-0000-4000-8000-000000000003/drive?$select=id,webUrl']: { id: 'drive-assets', webUrl: 'https://contoso.sharepoint.com/sites/Team/SiteAssets' },
  // no news on the site
  [NEWS_URL]: { value: [{ id: 'p-home', title: 'Home', promotionKind: 'page' }] }
});
const IMAGE_URL = BASE + '/drives/drive-assets/root:/SitePages/Home/copy-(1)-devis%20initech-232067.png:/content';

test('the home page of a team site: its sections, columns and blocks, the images inlined, what is not turned listed', () => {
  const lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: PNG });
  const content = JSON.parse(JSON.stringify(lib.gs_pageContent('token', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000000 })));
  assert.equal(content.page.name, 'Home.aspx');
  assert.equal(content.sections.length, 1);
  const [news, side] = content.sections[0].columns;
  assert.deepEqual([news.width, side.width], [8, 4]);
  // no news: nothing (as the readers of the site see it); the activity: the default library of the site
  assert.deepEqual(news.blocks, [{ type: 'activity', title: 'Activité du site', maxItems: 9, listId: 'aaaaaaaa-0000-4000-8000-000000000002', library: true,
    url: 'https://contoso.sharepoint.com/sites/Team/Shared%20Documents' }]);
  assert.equal(news.html, '<h3>Activité du site</h3>');
  assert.deepEqual(side.blocks.map((b) => b.type), ['links', 'list', 'text', 'image']);
  const links = side.blocks[0];
  assert.equal(links.title, 'Liens rapides');
  assert.equal(links.layout, 'List');
  assert.deepEqual(links.items.map((i) => [i.title, i.url, i.listId || null]), [
    ['contoso.sharepoint.com', 'https://contoso.sharepoint.com/sites/Team/Lists/Clients/AllItems.aspx', 'aaaaaaaa-0000-4000-8000-000000000001'],
    ['En savoir plus sur un site d’équipe', 'https://go.microsoft.com/fwlink/p/?linkid=827918', null]
  ]);
  assert.deepEqual(side.blocks[1], { type: 'list', title: 'Documents', listId: 'aaaaaaaa-0000-4000-8000-000000000002', library: true,
    url: 'https://contoso.sharepoint.com/sites/Team/Shared%20Documents' });
  assert.equal(side.blocks[3].status, 'embedded');
  assert.ok(side.blocks[3].src.startsWith('data:image/png;base64,'));
  assert.equal(content.imagesEmbedded, 1);
  assert.ok(side.html.includes('<h3>Liens rapides</h3>'));
  assert.ok(side.html.includes('<img src="data:image/png;base64,'));
  assert.deepEqual(content.skipped.map((s) => [s.type, s.title, s.reason]), [['news', 'Actualités', 'empty']]);
  // no token of the host of the site: no logo, nothing sent to it
  assert.deepEqual(content.logo, { status: 'unavailable', src: '', color: '' });
  assert.ok(!lib.calls.some((url) => url.includes('getsitelogo')));
  // the image read by Graph, the drive of the library that holds it, by path
  assert.ok(lib.calls.includes(IMAGE_URL));
});

test('the images: too large, too many, of another host (kept as they are), not inlined', () => {
  let lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: new Array(2000).fill(1) });
  let image = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000 }).sections[0].columns[1].blocks[3];
  assert.deepEqual([image.status, image.src], ['too_large', '']);
  lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: PNG });
  image = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 0, maxImageBytes: 1000 }).sections[0].columns[1].blocks[3];
  assert.equal(image.status, 'limit');
  image = lib.gs_pageContent('t', SITE, { embedImages: false, maxImages: 5, maxImageBytes: 1000 }).sections[0].columns[1].blocks[3];
  assert.equal(image.status, 'unavailable');
  // not an image (its first bytes): not inlined
  lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: [0x3c, 0x73, 0x76, 0x67] });
  image = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 5, maxImageBytes: 1000 }).sections[0].columns[1].blocks[3];
  assert.equal(image.status, 'unavailable');
  const page = fixture('home-page.json');
  page.canvasLayout.horizontalSections[0].columns[1].webparts[3].data.serverProcessedContent.imageSources = [{ key: 'imageSource', value: 'https://images.example.org/a.png' }];
  lib = load(answersOf(page), {});
  image = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 5, maxImageBytes: 1000 }).sections[0].columns[1].blocks[3];
  assert.deepEqual([image.status, image.src], ['external', 'https://images.example.org/a.png']);
  // a path out of the library: refused
  page.canvasLayout.horizontalSections[0].columns[1].webparts[3].data.serverProcessedContent.imageSources = [{ key: 'imageSource', value: '/sites/Team/SiteAssets/../../Other/x.png' }];
  lib = load(answersOf(page), {});
  image = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 5, maxImageBytes: 1000 }).sections[0].columns[1].blocks[3];
  assert.equal(image.status, 'unavailable');
});

test('the page chosen: by id, by name, else the home one; none: an error 404', () => {
  const lib = load(answersOf(fixture('home-page.json')), {});
  assert.equal(lib.gs_pageChoose('t', SITE, '', '').id, 'p-home');
  assert.equal(lib.gs_pageChoose('t', SITE, 'p-article', '').id, 'p-article');
  assert.equal(lib.gs_pageChoose('t', SITE, '', 'NEWS.aspx').id, 'p-article');
  assert.equal(lib.gs_pageChoose('t', SITE, 'nope', ''), null);
  assert.throws(() => lib.gs_pageContent('t', SITE, { pageId: 'nope' }), (e) => e.httpStatus === 404);
});

test('the HTML kept: the tags and attributes of the lists only, no script, event, style nor javascript URL', () => {
  const lib = load({}, {});
  const clean = (h) => lib.gs_cleanHtml(h, 'https://contoso.sharepoint.com');
  assert.equal(clean('<p style="color:red" class="x">Hi <strong>you</strong></p>'), '<p>Hi <strong>you</strong></p>');
  assert.equal(clean('<script>go()</script><p>a</p>'), '<p>a</p>');
  assert.equal(clean('<img src=x onerror=go()>'), '');
  assert.equal(clean('<p onclick="go()">a</p>'), '<p>a</p>');
  assert.equal(clean('<a href="javascript:go()">a</a>'), '<a>a</a>');
  assert.equal(clean('<a href="java&#115;cript:go()">a</a>'), '<a>a</a>');
  assert.equal(clean('<a href="jav&#x09;ascript:go()">a</a>'), '<a>a</a>');
  assert.equal(clean('<a href="/sites/Team/Lists/Clients">c</a>'), '<a href="https://contoso.sharepoint.com/sites/Team/Lists/Clients" target="_blank" rel="noopener noreferrer">c</a>');
  assert.equal(clean('<a href="//other.example/x">c</a>'), '<a>c</a>');
  assert.equal(clean('<iframe src="https://other.example"></iframe><svg onload=go()><p>a</p>'), '');
  assert.equal(clean('<p>a < b & c > d</p>'), '<p>a &lt; b &amp; c &gt; d</p>');
  assert.equal(clean('<p>&nbsp;&lt;x&gt;</p>'), '<p>&nbsp;&lt;x&gt;</p>');
  assert.equal(clean('<img src="data:image/svg+xml;base64,PHN2Zz4=">'), '');
  assert.equal(clean('<img src="data:image/png;base64,iVBO" alt="a&quot;b">'), '<img src="data:image/png;base64,iVBO" alt="a&quot;b">');
  assert.equal(clean('<div><p>unclosed'), '<div><p>unclosed</p></div>');
  assert.equal(clean('<p>a</div></p>'), '<p>a</p>');
  assert.equal(clean('<p title="x>y">a</p><!-- <script>x</script> -->'), '<p>a</p>');
  assert.equal(clean('<math><mi xlink:href="javascript:1">x</mi></math><p>b</p>'), '<p>b</p>');
  assert.equal(clean('<<script>script>go()</script>'), '&lt;');
  assert.equal(clean('<a href="mailto:a@b.example">m</a>'), '<a href="mailto:a@b.example" target="_blank" rel="noopener noreferrer">m</a>');
});

test('the news of the site: its posts, the newest first, their thumbnails inlined, as many as the web part shows', () => {
  const answers = answersOf(fixture('home-page.json'));
  answers[NEWS_URL] = { value: [
    { id: 'n1', title: 'Old', description: 'First post', promotionKind: 'newsPost', createdDateTime: '2026-01-01T10:00:00Z',
      webUrl: 'https://contoso.sharepoint.com/sites/Team/SitePages/Old.aspx' },
    { id: 'p-home', title: 'Home', promotionKind: 'page', createdDateTime: '2026-06-01T10:00:00Z' },
    { id: 'n2', title: 'New', description: '', promotionKind: 'newsPost', createdDateTime: '2026-03-01T10:00:00Z',
      webUrl: 'https://contoso.sharepoint.com/sites/Team/SitePages/New.aspx',
      thumbnailWebUrl: 'https://contoso.sharepoint.com/sites/Team/SiteAssets/SitePages/New/cover.png' },
    { id: 'n3', title: 'Bad', promotionKind: 'newsPost', createdDateTime: '2026-02-01T10:00:00Z', webUrl: 'javascript:go()' }
  ] };
  const cover = BASE + '/drives/drive-assets/root:/SitePages/New/cover.png:/content';
  const page = fixture('home-page.json');
  page.canvasLayout.horizontalSections[0].columns[0].webparts[0].data.properties = { layoutId: 'FeaturedNews', prefetchCount: 2 };
  answers[SITE_URL + '/pages/p-home/microsoft.graph.sitePage?$expand=canvasLayout'] = page;
  const lib = load(answers, { [IMAGE_URL]: PNG, [cover]: PNG });
  const content = JSON.parse(JSON.stringify(lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000000 })));
  const news = content.sections[0].columns[0].blocks[0];
  assert.equal(news.type, 'news');
  assert.equal(news.title, 'Actualités');
  assert.equal(news.layout, 'FeaturedNews');
  assert.deepEqual(news.items.map((i) => [i.title, i.url, i.image.status]), [
    ['New', 'https://contoso.sharepoint.com/sites/Team/SitePages/New.aspx', 'embedded'],
    ['Bad', '', 'unavailable']
  ]);
  assert.ok(news.items[0].image.src.startsWith('data:image/png;base64,'));
  assert.ok(!content.sections[0].columns[0].html.includes('javascript'));
  assert.deepEqual(content.skipped, []);
});

test('the news unreadable: left out, said in skipped', () => {
  const answers = answersOf(fixture('home-page.json'));
  delete answers[NEWS_URL];
  const lib = load(answers, { [IMAGE_URL]: PNG });
  const content = JSON.parse(JSON.stringify(lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000000 })));
  assert.deepEqual(content.sections[0].columns[0].blocks.map((b) => b.type), ['activity']);
  assert.deepEqual(content.skipped.map((s) => [s.type, s.reason]), [['news', 'unavailable']]);
});

test('the logo of the site: by its host with the token of that host, inlined with the color of its corner', () => {
  const logoUrl = "https://contoso.sharepoint.com/sites/Team/_api/siteiconmanager/getsitelogo?type=%271%27";
  let lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: PNG, [logoUrl]: PNG });
  lib.gs_pageCornerColor = () => '#CA5010';
  const tokens = [];
  const download = lib.gs_pageDownloadBytes;
  lib.gs_pageDownloadBytes = (url, token, max) => { tokens.push([url, token]); return download(url, token, max); };
  let content = JSON.parse(JSON.stringify(lib.gs_pageContent('graph-token', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000000, sharePointToken: 'sp-token' })));
  assert.equal(content.logo.status, 'embedded');
  assert.ok(content.logo.src.startsWith('data:image/png;base64,'));
  assert.equal(content.logo.color, '#CA5010');
  // the token of the host goes to the host only, the Graph token to Graph only
  assert.deepEqual(tokens.filter(([url]) => url === logoUrl).map(([, token]) => token), ['sp-token']);
  assert.ok(tokens.filter(([url]) => url !== logoUrl).every(([url, token]) => url.startsWith(BASE) && token === 'graph-token'));
  // not an image: not inlined
  lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: PNG, [logoUrl]: [0x3c, 0x68] });
  content = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000000, sharePointToken: 'sp' });
  assert.equal(content.logo.status, 'unavailable');
  // refused by the host: no logo, the page still read
  lib = load(answersOf(fixture('home-page.json')), { [IMAGE_URL]: PNG });
  content = lib.gs_pageContent('t', SITE, { embedImages: true, maxImages: 20, maxImageBytes: 1000000, sharePointToken: 'sp' });
  assert.deepEqual([content.logo.status, content.sections.length], ['unavailable', 1]);
});

test('the news of the site, whole: key, text without image nor link, the file of its thumbnail, the newest first', () => {
  const answers = answersOf(fixture('home-page.json'));
  const post = (id, title, created, thumbnail) => ({ id, title, description: 'Résumé de ' + title, promotionKind: 'newsPost', createdDateTime: created,
    webUrl: 'https://contoso.sharepoint.com/sites/Team/SitePages/' + title + '.aspx', thumbnailWebUrl: thumbnail });
  answers[NEWS_URL] = { value: [
    post('1a1b2c3d-4e5f-4061-8000-000000000001', 'Ancienne', '2026-01-01T10:00:00Z', 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'),
    { id: 'p-home', title: 'Home', promotionKind: 'page', createdDateTime: '2026-06-01T10:00:00Z' },
    post('0a1b2c3d-4e5f-4061-8000-000000000002', 'Nouvelle', '2026-03-01T10:00:00Z',
      'https://contoso.sharepoint.com/_layouts/15/getpreview.ashx?path=%2Fsites%2FTeam%2FSiteAssets%2FSitePages%2FNouvelle%2Fcover.png'),
    post('3a1b2c3d-4e5f-4061-8000-000000000003', 'Ailleurs', '2026-02-01T10:00:00Z', 'https://images.example.org/a.png')
  ] };
  answers[SITE_URL + '/pages/0a1b2c3d-4e5f-4061-8000-000000000002/microsoft.graph.sitePage?$expand=canvasLayout'] = { canvasLayout: { horizontalSections: [{ columns: [{ webparts: [
    { '@odata.type': '#microsoft.graph.textWebPart', innerHtml: '<p>Bonjour <a href="https://contoso.sharepoint.com/sites/Team/SitePages/X.aspx">la page</a></p><p><img src="/sites/Team/SiteAssets/a.png"></p><script>x</script>' },
    { '@odata.type': '#microsoft.graph.standardWebPart', webPartType: 'd1d91016-032f-456d-98a4-721247c305e8', data: { properties: {} } }] }] }] } };
  answers[BASE + '/drives/drive-assets/root:/SitePages/Nouvelle/cover.png:?$select=id,name,size,file'] = { id: 'item-cover', name: 'cover.png', size: 1234, file: {} };
  const lib = load(answers, {});
  const news = JSON.parse(JSON.stringify(lib.gs_siteNews('t', SITE, { withContent: true, maxItems: 50 })));
  assert.equal(news.count, 3);
  assert.deepEqual(news.items.map((i) => [i.title, i.key]), [['Nouvelle', 0x0a1b2c3d4e5f], ['Ailleurs', 0x3a1b2c3d4e5f], ['Ancienne', 0x1a1b2c3d4e5f]]);
  const [first, other, old] = news.items;
  assert.equal(first.pageId, '0a1b2c3d-4e5f-4061-8000-000000000002');
  assert.equal(first.description, 'Résumé de Nouvelle');
  // its text: no link to the site, no image, no script
  assert.equal(first.content, '<p>Bonjour la page</p><p></p>');
  // its thumbnail: the file of the library (the path of a preview of SharePoint); another host or the default one: none
  assert.deepEqual(first.image, { driveId: 'drive-assets', itemId: 'item-cover', name: 'cover.png', size: 1234 });
  assert.equal(other.image, null);
  assert.equal(old.image, null);
  // without the content: the list only, one read of Graph for the pages
  // a key met twice: the first post only
  answers[NEWS_URL].value.push(post('0a1b2c3d-4e5f-4061-8000-000000000009', 'Doublon', '2025-01-01T10:00:00Z', ''));
  const light = load(answers, {});
  const short = JSON.parse(JSON.stringify(light.gs_siteNews('t', SITE, { withContent: false, maxItems: 1 })));
  assert.deepEqual(short.items.map((i) => [i.title, i.content === undefined, i.image === undefined]), [['Nouvelle', true, true]]);
  assert.equal(short.count, 3);
  assert.ok(!JSON.stringify(light.gs_siteNews('t', SITE, { withContent: false, maxItems: 50 })).includes('Doublon'));
  assert.ok(!light.calls.some((url) => url.includes('$expand=canvasLayout') || url.includes('root:')));
});
