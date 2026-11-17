"""Self-check for sharepoint-mcp. Runs offline: no network, no SharePoint.

uv run test_server.py
"""

import contextlib
import io
import os
import tempfile
import zipfile
from pathlib import Path

import server as s


def test_path_escape_rejected():
    """Trust boundary: model-supplied paths must not escape the sync roots."""
    if not s.sync_roots():
        print('  skip: no sync root present')
        return
    for bad in ['../../etc/passwd', '/etc/passwd', '~/.ssh/id_rsa']:
        try:
            s._resolve_in_roots(bad)
        except ValueError:
            continue
        raise AssertionError(f'escape not blocked: {bad}')


def test_path_inside_root_allowed():
    roots = s.sync_roots()
    if not roots:
        print('  skip: no sync root present')
        return
    assert s._resolve_in_roots(str(roots[0])) == roots[0].resolve()


def test_hydrated_flag():
    """blocks==0 means placeholder; checking must not download."""

    class FakeStat:
        st_size = 21355
        st_blocks = 0

    class RealStat:
        st_size = 17149
        st_blocks = 40

    assert s._is_hydrated(FakeStat()) is False
    assert s._is_hydrated(RealStat()) is True


def _tmp_file(data: bytes) -> Path:
    """Write `data` to a fresh temp .xlsx; the caller unlinks it."""
    fd, name = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    out = Path(name)
    out.write_bytes(data)
    return out


def _synthetic_xlsx() -> Path:
    """Minimal OOXML workbook: one shared string, one numeric cell."""
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('xl/sharedStrings.xml', f'<sst xmlns="{ns}"><si><t>Viewer</t></si><si><t>Team</t></si></sst>')
        zf.writestr('xl/workbook.xml', f'<workbook xmlns="{ns}"><sheets><sheet name="Report"/></sheets></workbook>')
        zf.writestr(
            'xl/worksheets/sheet1.xml',
            f'<worksheet xmlns="{ns}"><sheetData>'
            '<row><c t="s"><v>0</v></c><c t="s"><v>1</v></c></row>'
            '<row><c><v>42</v></c><c t="inlineStr"><is><t>inline</t></is></c></row>'
            '</sheetData></worksheet>',
        )
    return _tmp_file(buf.getvalue())


def test_xlsx_extraction():
    p = _synthetic_xlsx()
    try:
        with zipfile.ZipFile(p) as zf:
            text = s._xlsx_text(zf, max_chars=10_000)
    finally:
        p.unlink(missing_ok=True)
    assert '### sheet: Report' in text, text
    assert 'Viewer\tTeam' in text, text  # shared strings resolved by index
    assert '42' in text, text  # numeric cell
    assert 'inline' in text, text  # inlineStr cell


def _multi_sheet_xlsx(n: int) -> Path:
    """Workbook with n sheets wired through a relationship table.

    Part filenames are emitted in an order that does NOT match document order,
    which is what breaks naive lexicographic pairing at 10+ sheets.
    """
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg = 'http://schemas.openxmlformats.org/package/2006/relationships'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        sheets = ''.join(f'<sheet name="S{i}" r:id="rId{i}"/>' for i in range(1, n + 1))
        zf.writestr('xl/workbook.xml', f'<workbook xmlns="{main}" xmlns:r="{rel}"><sheets>{sheets}</sheets></workbook>')
        rels = ''.join(f'<Relationship Id="rId{i}" Target="worksheets/sheet{i}.xml"/>' for i in range(1, n + 1))
        zf.writestr('xl/_rels/workbook.xml.rels', f'<Relationships xmlns="{pkg}">{rels}</Relationships>')
        for i in range(1, n + 1):
            zf.writestr(
                f'xl/worksheets/sheet{i}.xml',
                f'<worksheet xmlns="{main}"><sheetData><row>'
                f'<c t="inlineStr"><is><t>MARKER{i}</t></is></c>'
                '</row></sheetData></worksheet>',
            )
    return _tmp_file(buf.getvalue())


def test_xlsx_sheet_labels_match_content():
    """Label N must sit above MARKER N, even past 9 sheets."""
    p = _multi_sheet_xlsx(11)
    try:
        with zipfile.ZipFile(p) as zf:
            text = s._xlsx_text(zf, max_chars=100_000)
    finally:
        p.unlink(missing_ok=True)
    for i in range(1, 12):
        block = text.split(f'### sheet: S{i}\n')
        assert len(block) == 2, f'missing label S{i}'
        assert block[1].startswith(f'MARKER{i}'), f'S{i} labelled over wrong content: {block[1][:24]!r}'


def test_xlsx_fallback_sorts_numerically():
    """Without a rel table, sheet2 must precede sheet10."""
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        for i in (1, 2, 10):
            zf.writestr(f'xl/worksheets/sheet{i}.xml', f'<worksheet xmlns="{main}"/>')
    p = _tmp_file(buf.getvalue())
    try:
        with zipfile.ZipFile(p) as zf:
            order = [name for name, _ in s._xlsx_sheet_parts(zf)]
    finally:
        p.unlink(missing_ok=True)
    assert order == ['sheet1', 'sheet2', 'sheet10'], order


@contextlib.contextmanager
def fake_root():
    """A throwaway sync root.

    Never write into the real synced library -- anything created there uploads
    straight to corporate SharePoint.
    """
    with tempfile.TemporaryDirectory() as d:
        root = Path(d) / 'OneDrive-SharedLibraries-Test'
        root.mkdir()
        original = s.sync_roots
        s.sync_roots = lambda: [root]
        try:
            yield root
        finally:
            s.sync_roots = original


def test_no_html_unescaping_of_file_content():
    """Ampersand entities in plain text must survive verbatim."""
    with fake_root() as root:
        f = root / 'amp.csv'
        f.write_text('a,b\nx=1&amp;y=2,&amp;amp;\n')
        r = s.sp_read(str(f))
        assert 'x=1&amp;y=2' in r['content'], r['content']
        assert '&amp;amp;' in r['content'], r['content']


def test_office_text_not_double_unescaped():
    """A cell literally showing &lt;tag&gt; must not become <tag>."""
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    with fake_root() as root:
        f = root / 'esc.xlsx'
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w') as zf:
            zf.writestr(
                'xl/worksheets/sheet1.xml',
                f'<worksheet xmlns="{main}"><sheetData><row>'
                '<c t="inlineStr"><is><t>&amp;lt;tag&amp;gt;</t></is></c>'
                '</row></sheetData></worksheet>',
            )
        f.write_bytes(buf.getvalue())
        content = s.sp_read(str(f))['content']
        assert '&lt;tag&gt;' in content, content
        assert '<tag>' not in content, content


def test_size_gate_blocks_huge_files():
    assert s.MAX_READ_BYTES > 0
    with fake_root() as root:
        f = root / 'big.txt'
        f.write_text('x' * 100)
        orig = s.MAX_READ_BYTES
        try:
            s.MAX_READ_BYTES = 10
            assert 'over the' in s.sp_read(str(f)).get('error', '')
            s.MAX_READ_BYTES = orig
            assert 'content' in s.sp_read(str(f))
        finally:
            s.MAX_READ_BYTES = orig


def test_sp_find_bad_subdir_returns_error():
    with fake_root():
        r = s.sp_find('*.xlsx', subdir='/etc')
        assert 'error' in r, r


def test_sp_find_returns_newest_when_truncated():
    """Limit must yield the newest matches, not a walk-order slice."""
    with fake_root() as root:
        for i in range(6):
            f = root / f'f{i}.txt'
            f.write_text('x')
            os.utime(f, (1_700_000_000 + i * 86400,) * 2)
        r = s.sp_find('*.txt', limit=2)
        assert r['truncated'] is True, r
        assert r['total_matched'] == 6, r
        assert [h['name'] for h in r['results']] == ['f5.txt', 'f4.txt'], r['results']


def test_hydrated_treats_empty_file_as_local():
    class EmptyReal:
        st_size = 0
        st_blocks = 0

    class Placeholder:
        st_size = 21355
        st_blocks = 0

    assert s._is_hydrated(EmptyReal()) is True
    assert s._is_hydrated(Placeholder()) is False


def test_xml_text_survives_bad_xml():
    assert s._xml_text(b'<not xml') == ''
    assert 'hello' in s._xml_text(b'<a><b>hello</b></a>')


def test_read_only_by_default():
    """Writes are opt-in: nothing write-ish is advertised unless SHAREPOINT_ENABLE_WRITES is set."""
    assert s.ENABLE_WRITES is False
    assert s.REGISTERED_WRITE_TOOLS == []
    # The write functions exist so they stay testable, but must not be registered.
    assert callable(s.sp_list_item_update)
    # Create and delete must never exist, flag or not.
    assert not hasattr(s, 'sp_list_item_create')
    assert not hasattr(s, 'sp_list_item_delete')
    assert [t.__name__ for t, _ in s.WRITE_TOOLS] == ['sp_list_item_update']
    forbidden = {'mail_send', 'mail_reply', 'event_create', 'sp_write', 'mail_move'}
    names = {a for a in dir(s) if callable(getattr(s, a)) and not a.startswith('_')}
    assert not (names & forbidden), names & forbidden


# --- SharePoint list tools (Microsoft Graph) ------------------------------


def test_site_path_forms():
    """URLs, host:path, and composite ids all resolve to a Graph path segment."""
    assert s._site_path('https://contoso.sharepoint.com/sites/Foo') == 'contoso.sharepoint.com:/sites/Foo'
    assert s._site_path('https://contoso.sharepoint.com/sites/Foo/') == 'contoso.sharepoint.com:/sites/Foo'
    assert s._site_path('contoso.sharepoint.com:/sites/Foo') == 'contoso.sharepoint.com:/sites/Foo'
    composite = 'contoso.sharepoint.com,0dcf,abc'
    assert s._site_path(composite) == composite


def test_resolve_site_id_passthrough():
    """A composite id must not trigger a Graph round-trip."""
    composite = 'contoso.sharepoint.com,0dcf,abc'
    assert s._resolve_site_id(composite) == composite


def test_list_items_paging_and_limit():
    """Paging is followed and `limit` caps the result with a truncation flag."""
    pages = [
        {
            'value': [{'id': '1', 'fields': {'t': 'a'}}, {'id': '2', 'fields': {'t': 'b'}}],
            '@odata.nextLink': 'https://graph.microsoft.com/next',
        },
        {'value': [{'id': '3', 'fields': {'t': 'c'}}]},
    ]
    calls = []
    orig = s._graph_get
    s._graph_get = lambda url, headers=None: (calls.append(url), pages.pop(0))[1]
    try:
        r = s.sp_list_items('contoso.sharepoint.com,aaa,bbb', 'L', limit=2)
    finally:
        s._graph_get = orig
    assert r['count'] == 2, r
    assert r['truncated'] is True, r
    assert [i['id'] for i in r['items']] == ['1', '2'], r
    assert len(calls) == 2, calls  # a second page is fetched to detect the overflow


def test_list_items_filter_prefix_and_header():
    """filter_query gains the fields/ prefix; the Prefer header is always sent."""
    seen = {}

    def fake(url, headers=None):
        seen['url'] = url
        seen['headers'] = headers
        return {'value': []}

    orig = s._graph_get
    s._graph_get = fake
    try:
        s.sp_list_items('contoso.sharepoint.com,aaa,bbb', 'Projects', filter_query="status eq 'In Progress'")
    finally:
        s._graph_get = orig
    assert 'Projects' in seen['url'], seen['url']
    assert "$filter=fields/status%20eq%20'In%20Progress'" in seen['url'], seen['url']
    assert seen['headers'] == {'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly'}, seen['headers']


def test_graph_error_surfaces_as_dict():
    """A Graph failure must become {'error': ...}, never a raised exception."""

    def boom(url, headers=None):
        raise s.GraphError('boom')

    orig = s._graph_get
    s._graph_get = boom
    try:
        assert s.sp_lists('site') == {'error': 'boom'}
        assert s.sp_list_items('site', 'L') == {'error': 'boom'}
    finally:
        s._graph_get = orig


def test_az_bin_missing_raises_graph_error():
    """An absent `az` is reported as a GraphError, not a bare OSError."""
    orig = s.shutil.which
    s.shutil.which = lambda _name: None
    try:
        try:
            s._az_bin()
        except s.GraphError:
            return
        raise AssertionError('expected GraphError when az is absent')
    finally:
        s.shutil.which = orig


# --- SharePoint list writes (opt-in) --------------------------------------


def test_list_item_update_patches_only_given_fields():
    seen = {}

    calls = []

    def fake(url, method='GET', body=None, headers=None, timeout=60):
        calls.append(method)
        seen.update(url=url, method=method, body=body)
        return {'Title': 'old'} if method == 'GET' else {'Title': 'new'}

    orig = s._graph_request
    s._graph_request = fake
    try:
        r = s.sp_list_item_update('contoso.sharepoint.com,a,b', 'Projects', '3', {'Title': 'new'})
    finally:
        s._graph_request = orig
    assert seen['method'] == 'PATCH', seen
    assert seen['body'] == {'Title': 'new'}, seen
    assert '/items/3/fields' in seen['url'], seen
    assert 'Projects' in seen['url'], seen
    assert calls == ['GET', 'PATCH'], calls
    assert r == {'id': '3', 'changes': {'Title': {'before': 'old', 'after': 'new'}}}, r


def test_writes_reject_empty_fields():
    assert 'error' in s.sp_list_item_update('contoso.sharepoint.com,a,b', 'L', '3', {})


def test_write_error_surfaces_as_dict():
    def boom(url, method='GET', body=None, headers=None, timeout=60):
        raise s.GraphError('nope')

    orig = s._graph_request
    s._graph_request = boom
    try:
        assert s.sp_list_item_update('contoso.sharepoint.com,a,b', 'L', '3', {'T': 1}) == {'error': 'nope'}
    finally:
        s._graph_request = orig


def test_hydrated_windows_attributes():
    """Windows has no st_blocks; placeholder status comes from file attributes."""

    class Placeholder:
        st_size = 5000
        st_file_attributes = s.WIN_RECALL_ON_DATA_ACCESS | 0x20

    class Offline:
        st_size = 5000
        st_file_attributes = s.WIN_OFFLINE

    class Local:
        st_size = 5000
        st_file_attributes = 0x20  # FILE_ATTRIBUTE_ARCHIVE only

    assert s._is_hydrated(Placeholder()) is False
    assert s._is_hydrated(Offline()) is False
    assert s._is_hydrated(Local()) is True


def test_az_cmd_runs_bundled_python():
    """On Windows, az.cmd is bypassed so cmd.exe never re-parses arguments."""
    with tempfile.TemporaryDirectory() as d:
        cli = Path(d) / 'CLI2'
        (cli / 'wbin').mkdir(parents=True)
        (cli / 'python.exe').write_text('')
        az_cmd = cli / 'wbin' / 'az.cmd'
        orig = s.shutil.which
        s.shutil.which = lambda _name: str(az_cmd)
        try:
            assert s._az_bin() == [str(cli / 'python.exe'), '-IBm', 'azure.cli']
            (cli / 'python.exe').unlink()
            try:
                s._az_bin()
            except s.GraphError:
                pass
            else:
                raise AssertionError('expected GraphError without bundled python.exe')
        finally:
            s.shutil.which = orig


def test_windows_sync_roots_from_registry():
    """Roots come from OneDrive's registry: org folders plus the work OneDrive, never consumer OneDrive."""
    import sys
    import types

    with tempfile.TemporaryDirectory() as d:
        org = Path(d) / 'Contoso'
        (org / 'Team - Documents').mkdir(parents=True)
        personal = Path(d) / 'OneDrive - Contoso'
        (personal / 'Shortcut Library').mkdir(parents=True)
        (Path(d) / 'OneDrive').mkdir()
        entries = {
            'a': {
                'MountPoint': str(org / 'Team - Documents'),
                'UrlNamespace': 'https://contoso.sharepoint.com/sites/team/',
            },
            'b': {'MountPoint': str(personal), 'UrlNamespace': 'https://contoso-my.sharepoint.com/personal/me/'},
            'c': {},  # an entry without values is ignored
            'd': {'MountPoint': str(Path(d) / 'OneDrive'), 'UrlNamespace': 'https://d.docs.live.net/abc123/'},
        }

        class Key:
            def __init__(self, name=None):
                self.name = name

            def __enter__(self):
                return self

            def __exit__(self, *exc):
                return False

        fake = types.SimpleNamespace(
            HKEY_CURRENT_USER=object(),
            OpenKey=lambda parent, sub: Key(sub if isinstance(parent, Key) else None),
            QueryInfoKey=lambda _key: (len(entries), 0, 0),
            EnumKey=lambda _key, i: list(entries)[i],
        )

        def query(key, value):
            try:
                return (entries[key.name][value], 1)
            except KeyError:
                raise OSError(value) from None

        fake.QueryValueEx = query
        sys.modules['winreg'] = fake
        try:
            assert s._windows_sync_roots() == sorted([org, personal])
        finally:
            del sys.modules['winreg']


def test_mac_sync_roots_include_work_onedrive():
    """Shortcut-added libraries live in the work OneDrive, so it is a root too."""
    if s.sys.platform == 'win32':
        return
    with tempfile.TemporaryDirectory() as d:
        cloud = Path(d) / 'Library' / 'CloudStorage'
        for name in ('OneDrive-SharedLibraries-Contoso', 'OneDrive-Contoso', 'OneDrive-Personal', 'GoogleDrive-me'):
            (cloud / name).mkdir(parents=True)
        orig = s.Path.home
        s.Path.home = classmethod(lambda _cls: Path(d))
        try:
            names = [p.name for p in s.sync_roots()]
        finally:
            s.Path.home = orig
        assert names == ['OneDrive-Contoso', 'OneDrive-SharedLibraries-Contoso'], names


def test_my_sites_batches_and_reports_groups_without_site():
    """Groups are read once, sites come back in one $batch, 404s are listed."""
    calls = []

    def fake(url, method='GET', body=None, headers=None, timeout=60):
        calls.append((method, url))
        if 'memberOf' in url:
            assert headers == {'ConsistencyLevel': 'eventual'}
            return {'value': [{'id': 'g1', 'displayName': 'Beta Team'}, {'id': 'g2', 'displayName': 'Alpha'}]}
        assert url.endswith('/$batch')
        assert method == 'POST'
        site = {'id': 'h,1,2', 'displayName': 'Beta Team', 'webUrl': 'https://x/b'}
        return {
            'responses': [
                {'id': r['id'], 'status': 200, 'body': site}
                if r['url'].startswith('/groups/g1/')
                else {'id': r['id'], 'status': 404, 'body': {'error': {}}}
                for r in reversed(body['requests'])  # Graph may answer out of order
            ]
        }

    orig = s._graph_request
    s._graph_request = fake
    try:
        r = s.sp_my_sites()
        assert r['sites'] == [{'name': 'Beta Team', 'site_id': 'h,1,2', 'web_url': 'https://x/b'}], r
        assert r['without_site'] == ['Alpha']
        assert len(calls) == 2
        assert s.sp_my_sites(query='alp')['sites'] == []  # only Alpha matches, and it has no site
    finally:
        s._graph_request = orig


def test_list_schema_types_choices_and_hidden():
    cols = [
        {
            'name': 'Status',
            'displayName': 'Status',
            'choice': {'choices': ['New', 'Done'], 'displayAs': 'dropDownMenu'},
        },
        {'name': 'Owner', 'displayName': 'Owner', 'personOrGroup': {'allowMultipleSelection': True}, 'indexed': True},
        {'name': 'ID', 'displayName': 'ID', 'readOnly': True},
        {'name': '_Hidden', 'hidden': True, 'text': {}},
    ]
    orig = s._graph_get
    s._graph_get = lambda url, headers=None, timeout=60: {'value': cols}
    try:
        r = s.sp_list_schema('h,1,2', 'Projects')
        by = {c['name']: c for c in r['columns']}
        assert set(by) == {'Status', 'Owner', 'ID'}
        assert by['Status']['type'] == 'choice'
        assert by['Status']['choices'] == ['New', 'Done']
        assert by['Owner']['type'] == 'personOrGroup'
        assert by['Owner']['multiple']
        assert by['Owner']['indexed']
        assert by['ID']['type'] == 'other'
        assert by['ID']['read_only']
        assert s.sp_list_schema('h,1,2', 'Projects', include_hidden=True)['count'] == 4
    finally:
        s._graph_get = orig


def test_recent_changes_query_and_mapping():
    seen = []

    def fake(url, headers=None, timeout=60):
        seen.append(url)
        if url.endswith('/drive/list?$select=id'):
            return {'id': 'L1'}
        return {
            'value': [
                {
                    'id': '7',
                    'webUrl': 'https://x/f.docx',
                    'lastModifiedBy': {'user': {'displayName': 'Ann'}},
                    'fields': {
                        'FileLeafRef': 'f.docx',
                        'FileRef': '/sites/x/f.docx',
                        'Modified': '2026-11-12T10:00:00Z',
                        'ContentType': 'Document',
                    },
                },
                {'id': '8', 'fields': {'Title': 'Row', 'Modified': '2026-11-11T10:00:00Z', 'ContentType': 'Folder'}},
            ],
            '@odata.nextLink': 'https://graph/next',
        }

    orig = s._graph_get
    s._graph_get = fake
    try:
        r = s.sp_recent_changes('h,1,2', since='2026-11-10', limit=2)
        assert '/lists/L1/items?' in seen[1]
        assert "fields/Modified%20ge%20'2026-11-10T00:00:00Z'" in seen[1]
        assert '$orderby=fields/Modified%20desc' in seen[1]
        assert r['items'][0] == {
            'id': '7', 'name': 'f.docx', 'path': '/sites/x/f.docx', 'modified': '2026-11-12T10:00:00Z',
            'modified_by': 'Ann', 'is_folder': False, 'web_url': 'https://x/f.docx',
        }  # fmt: skip
        assert r['items'][1]['name'] == 'Row'
        assert r['items'][1]['is_folder']
        assert len(seen) == 3
        assert r['truncated']
        r = s.sp_recent_changes('h,1,2', since='2026-11-10', limit=1)
        assert r['truncated']
        assert r['count'] == 1
    finally:
        s._graph_get = orig


def test_recent_changes_rejects_odata_injection():
    for bad in ("2026-11-10' or 1 eq 1", 'last week', '2026-11-10T10:00:00+02:00'):
        assert 'error' in s.sp_recent_changes('h,1,2', since=bad), bad


def _run(t) -> bool:
    """Run one test, print its outcome, and return whether it passed."""
    try:
        t()
    except Exception as e:
        print(f'FAIL {t.__name__}: {type(e).__name__}: {e}')
        return False
    print(f'ok   {t.__name__}')
    return True


if __name__ == '__main__':
    tests = [v for k, v in sorted(globals().items()) if k.startswith('test_')]
    failed = sum(not _run(t) for t in tests)
    print(f'\n{len(tests) - failed}/{len(tests)} passed')
    raise SystemExit(1 if failed else 0)
