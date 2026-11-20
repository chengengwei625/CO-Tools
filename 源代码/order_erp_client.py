"""Submit order colour lists to the two i9SCM colour libraries independently."""
import base64
import json
import os
import uuid
from contextlib import contextmanager
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from tempfile import NamedTemporaryFile
from http.cookiejar import CookieJar
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import HTTPRedirectHandler, HTTPCookieProcessor, Request, build_opener


ERP_URL = 'http://121.11.144.4:50082'
LIBRARIES = (
    ('物料颜色', 'material-color', 'MatColor', 'getMatColorSeries'),
    ('成品颜色', 'style-color', 'ProdColor', 'getColorStyleSeries'),
)


class ErpError(Exception):
    pass


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def _text(value):
    return str(value if value is not None else '').strip()


def _key(value):
    return ' '.join(_text(value).split()).casefold()


def _form(fields):
    boundary = 'OrderColour' + uuid.uuid4().hex
    body = ''.join(f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"'
                   f'\r\n\r\n{value}\r\n' for key, value in fields.items())
    return (body + f'--{boundary}--\r\n').encode('utf-8'), f'multipart/form-data; boundary={boundary}'


# Exclude query metadata such as row numbers, which can change between reads.
UNDO_COLOUR_FIELDS = (
    'ColorId', 'ColorName', 'SeriesId', 'SeriesNm', 'ColorEng', 'Operator', 'EditDate',
    'ID',
    'PantoneId', 'ColorShow', 'ColorO', 'ColorCW', 'Des', 'PantonePic', 'PicturePath',
)


def _record_fingerprint(row):
    return [_text(row.get(field)) for field in UNDO_COLOUR_FIELDS]


class OrderErpClient:
    def __init__(self, undo_directory=None):
        self._undo_directory = Path(undo_directory) if undo_directory is not None else None
        self._undo_path = None
        self._batch = self._empty_batch()
        self._opener = build_opener(NoRedirect(), HTTPCookieProcessor(CookieJar()))
        self._token = ''
        self.display_name = ''

    @property
    def logged_in(self):
        return bool(self._token)

    def _request(self, path, *, params=None, fields=None):
        headers = {'Accept': 'application/json', 'language': 'zh', 'systemtype': '1',
                   'channelid': '', 'Referer': ERP_URL + '/i9scm/'}
        if self._token:
            headers['token'] = self._token
        body = None
        if fields is not None:
            body, headers['Content-Type'] = _form(fields)
        url = ERP_URL + '/index.php/' + path
        if params:
            url += '?' + urlencode(params)
        try:
            with self._opener.open(Request(url, data=body, headers=headers), timeout=20) as response:
                result = json.load(response)
        except HTTPError as ex:
            if ex.code in (401, 403):
                self._token = ''
                raise ErpError('罗斯系统登录已失效或没有操作权限。') from ex
            raise ErpError(f'罗斯系统请求失败（HTTP {ex.code}）。') from ex
        except (URLError, OSError) as ex:
            raise ErpError('无法连接罗斯系统，请检查网络后重试。') from ex
        except ValueError as ex:
            raise ErpError('罗斯系统返回了无法识别的内容。') from ex
        if not isinstance(result, dict) or str(result.get('code')) != '1':
            message = _text(result.get('msg')) if isinstance(result, dict) else ''
            if any(text in message for text in ('未登录', '请先登录', '重新登录', '登录超时', '登录失效', '登录已失效', 'token失效')):
                self._token = ''
            raise ErpError(message or '罗斯系统未完成请求。')
        return result

    def login(self, username, password):
        if not username.strip() or not password:
            raise ErpError('请输入账号和密码。')
        self._token = ''
        encoded_password = base64.b64encode(password.encode('utf-8')).decode('ascii')
        try:
            result = self._request('adminscm/UserAdmin/loginnew', fields={
                'login_type': '0', 'channelId': '1', 'remember': '0', 'encrypt': '1',
                'username': username.strip(), 'password': encoded_password,
            })
        except ErpError as ex:
            message = str(ex).replace(encoded_password, '***').replace(password, '***')
            raise ErpError(message) from None
        user = result.get('user')
        token = result.get('token')
        if not token or not isinstance(user, dict):
            raise ErpError('登录响应缺少会话信息。')
        self._token = token
        self.display_name = _text(user.get('Usrnm') or user.get('user_name'))
        self._batch = self._empty_batch()
        if self._undo_directory is not None:
            account = sha256(_key(username).encode('utf-8')).hexdigest()
            self._undo_path = self._undo_directory / f'order_erp_undo_{account}.json'

    @staticmethod
    def _empty_batch():
        return {'started_at': '', 'rows': []}

    @staticmethod
    def _valid_snapshot(snapshot):
        return (isinstance(snapshot, list) and len(snapshot) == len(UNDO_COLOUR_FIELDS)
                and all(isinstance(value, str) for value in snapshot) and bool(snapshot[0]))

    def import_result(self):
        if self._undo_path is None:
            return self._batch
        try:
            batch = json.loads(self._undo_path.read_text(encoding='utf-8'))
        except FileNotFoundError:
            return self._batch
        except (OSError, ValueError) as ex:
            raise ErpError('无法读取本次 ERP 导入记录。') from ex

        if isinstance(batch, list):
            batch = self._legacy_batch(batch)
        if (not isinstance(batch, dict) or not isinstance(batch.get('started_at'), str)
                or not isinstance(batch.get('rows'), list) or len(batch['rows']) > 4):
            raise ErpError('本次 ERP 导入记录格式不正确。')
        modules = {module for _, module, _, _ in LIBRARIES}
        for row in batch['rows']:
            if not isinstance(row, dict) or not all(isinstance(row.get(key), str) for key in (
                    'module', 'code', 'name', 'category', 'english', 'status', 'detail')):
                raise ErpError('本次 ERP 导入记录格式不正确。')
            if row['module'] not in modules or row['status'] not in ('added', 'existing', 'failed'):
                raise ErpError('本次 ERP 导入记录格式不正确。')
            if row['status'] == 'added':
                snapshot = row.get('snapshot')
                if not self._valid_snapshot(snapshot) or _key(row['code']) != _key(snapshot[0]):
                    raise ErpError('本次 ERP 导入记录缺少成功新增的核对资料。')
        return batch

    def _legacy_batch(self, records):
        # Earlier versions only saved verified new records, without the other outcomes.
        rows = []
        if len(records) > 4:
            raise ErpError('本次 ERP 导入记录格式不正确。')
        for record in records:
            if not isinstance(record, list) or len(record) != 2 or not self._valid_snapshot(record[1]):
                raise ErpError('本次 ERP 导入记录格式不正确。')
            module, snapshot = record
            code, name, series_id, series, english = snapshot[:5]
            rows.append({'module': module, 'code': code, 'name': name, 'english': english,
                         'category': f'{series}({series_id})', 'status': 'added', 'snapshot': snapshot,
                         'detail': '旧版仅保存已核实新增的记录。'})
        return {'started_at': '', 'rows': rows}

    @staticmethod
    def _undo_entries(batch):
        return [row for row in batch['rows'] if row['status'] == 'added' and not row.get('undo')]

    def undo_records(self):
        return [[row['module'], row['snapshot']] for row in self._undo_entries(self.import_result())]

    def _save_batch(self, batch):
        if self._undo_path is not None:
            temporary = None
            try:
                self._undo_path.parent.mkdir(parents=True, exist_ok=True)
                with NamedTemporaryFile(mode='w', encoding='utf-8', dir=self._undo_path.parent,
                                        prefix='erp_undo_', suffix='.tmp', delete=False) as stream:
                    temporary = Path(stream.name)
                    json.dump(batch, stream, ensure_ascii=False)
                os.replace(temporary, self._undo_path)
            except OSError as ex:
                raise ErpError('无法保存本次 ERP 导入记录。') from ex
            finally:
                if temporary is not None:
                    try:
                        temporary.unlink(missing_ok=True)
                    except OSError:
                        pass
        self._batch = batch

    @contextmanager
    def _lock_undo_records(self):
        if self._undo_path is None:
            yield
            return
        try:
            self._undo_path.parent.mkdir(parents=True, exist_ok=True)
            lock = self._undo_path.with_suffix('.lock').open('a+b')
        except OSError as ex:
            raise ErpError('无法打开本次 ERP 导入记录。') from ex
        with lock:
            try:
                if os.name == 'nt':
                    from msvcrt import LK_NBLCK, locking
                    lock.seek(0)
                    locking(lock.fileno(), LK_NBLCK, 1)
                else:
                    from fcntl import LOCK_EX, LOCK_NB, flock
                    flock(lock.fileno(), LOCK_EX | LOCK_NB)
            except OSError as ex:
                raise ErpError('另一个窗口正在导入或删除，请稍后重试。') from ex
            # Closing the file releases the lock, including when a request fails.
            yield

    def _series(self, module, controller, series_endpoint):
        result = self._request(f'baseInformation/{controller}/{series_endpoint}',
                               params={'ModuleFunctionCode': module})
        series = result.get('data')
        if not isinstance(series, list) or any(
                not isinstance(row, dict) or not {'SeriesID', 'SeriesNm'} <= row.keys() for row in series):
            raise ErpError('罗斯系统返回的颜色分类格式不正确。')
        return {_text(row['SeriesID']): _text(row['SeriesNm']) for row in series}

    def _existing(self, module, controller, code):
        result = self._request(f'baseInformation/{controller}/base{controller}', params={
            'ModuleFunctionCode': module, 'ColorId': code, 'ColorName': '', 'SeriesNm': '',
            'SeriesId': '', 'ColorO': '', 'PantoneId': '', 'ColorEng': '', 'ColorShow': '',
            'page': 1, 'page_count': 1000,
        })
        rows = result.get('data', result.get('List'))
        if not isinstance(rows, list) or any(not isinstance(row, dict) or 'ColorId' not in row for row in rows):
            raise ErpError('罗斯系统返回的颜色清单格式不正确。')
        matches = [row for row in rows if _key(row.get('ColorId')) == _key(code)]
        if len(matches) > 1:
            raise ErpError(f'罗斯系统存在重复色号 {code}，请先核对。')
        try:
            count = int(result.get('Count', len(rows)))
        except (TypeError, ValueError) as ex:
            raise ErpError('罗斯系统返回的颜色记录数格式不正确。') from ex
        if not matches and count > len(rows):
            raise ErpError(f'罗斯系统色号 {code} 的查询结果不完整。')
        return matches[0] if matches else None

    @staticmethod
    def _colour_differences(existing, name, series_id, english):
        fields = (('ColorName', name, '颜色名称'), ('SeriesId', series_id, '颜色分类'),
                  ('ColorEng', english, '英文名称'))
        return [label for field, expected, label in fields if _key(existing.get(field)) != _key(expected)]

    @classmethod
    def _existing_detail(cls, existing, name, series_id, english):
        differences = cls._colour_differences(existing, name, series_id, english)
        detail = 'ERP 已有相同色号，已跳过。'
        if differences:
            detail += f'{"、".join(differences)}与清单不同。'
        return detail

    def import_colours(self, material_rows, finished_rows):
        if not self.logged_in:
            raise ErpError('请先登录罗斯系统。')
        with self._lock_undo_records():
            return self._import_colours(material_rows, finished_rows)

    def _import_colours(self, material_rows, finished_rows):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        batch = {'started_at': now, 'rows': []}
        groups = []
        for (_, module, controller, endpoint), candidates in zip(LIBRARIES, (material_rows, finished_rows)):
            if len(candidates) > 2:
                raise ErpError('每个颜色库只导入清单前两条。')
            rows = [{'module': module, 'code': code, 'name': name, 'category': category,
                     'english': english, 'status': 'failed', 'detail': '未提交或尚未核实。'}
                    for code, name, category, english in candidates]
            batch['rows'].extend(rows)
            groups.append((module, controller, endpoint, rows))
        self._save_batch(batch)

        pending = []
        preflight_failed = False
        for module, controller, endpoint, rows in groups:
            if not rows:
                continue
            try:
                series = self._series(module, controller, endpoint)
            except ErpError as ex:
                preflight_failed = True
                for row in rows:
                    row['detail'] = str(ex)
                continue
            codes = [_key(row['code']) for row in rows]
            for row, key in zip(rows, codes):
                code, name, category, english = (row[k] for k in ('code', 'name', 'category', 'english'))
                series_name, sep, series_id = category.rstrip(')').rpartition('(')
                try:
                    if not key or codes.count(key) > 1:
                        raise ErpError('待导入色号为空或与本清单中的其他记录重复。')
                    existing = self._existing(module, controller, code)
                    if existing is not None:
                        row.update(status='existing', detail=self._existing_detail(existing, name, series_id, english))
                    elif not name or not sep or series.get(series_id) != series_name:
                        raise ErpError('缺少颜色名称或分类无效，请先检查清单。')
                    else:
                        pending.append((controller, row, series_id))
                except ErpError as ex:
                    preflight_failed = True
                    row['detail'] = str(ex)
            self._save_batch(batch)

        if preflight_failed:
            for _, row, _ in pending:
                row['detail'] = '本批存在核对异常，未提交此记录。'
            self._save_batch(batch)
            return batch

        for index, (controller, row, series_id) in enumerate(pending):
            module, code, name, english = (row[k] for k in ('module', 'code', 'name', 'english'))
            try:
                existing = self._existing(module, controller, code)
                if existing is not None:
                    row.update(status='existing', detail=self._existing_detail(existing, name, series_id, english))
                else:
                    payload = {'EditState': '1', 'ColorId': code, 'ColorName': name, 'SeriesId': series_id,
                               'ColorEng': english, 'Operator': self.display_name, 'EditDate': now, 'Action': 0,
                               'PantoneId': '', 'ColorShow': '', 'ColorO': '', 'Des': '', 'PantonePic': '',
                               'PicturePath': ''}
                    self._request(f'baseInformation/{controller}/save{controller}', fields={
                        'ModuleFunctionCode': module, 'data': json.dumps(payload, ensure_ascii=False)})
                    created = self._existing(module, controller, code)
                    if created is None or self._colour_differences(created, name, series_id, english):
                        raise ErpError('保存后无法核实，请先在 ERP 核对。')
                    row.update(status='added', detail='已核实新增。', snapshot=_record_fingerprint(created))
            except ErpError as ex:
                row['detail'] = str(ex)
                for _, skipped, _ in pending[index + 1:]:
                    skipped['detail'] = '前面的记录导入异常，未提交此记录。'
                self._save_batch(batch)
                break
            self._save_batch(batch)
        return batch

    def undo_import(self, records):
        if not self.logged_in:
            raise ErpError('请先登录罗斯系统。')
        with self._lock_undo_records():
            return self._undo_import(records)

    def _undo_import(self, records):
        batch = self.import_result()
        entries = self._undo_entries(batch)
        if [[row['module'], row['snapshot']] for row in entries] != records:
            raise ErpError('本次导入记录已变化，请重新点击删除。')

        targets = {module: (label, controller) for label, module, controller, _ in LIBRARIES}
        results = {label: {'deleted': [], 'absent': [], 'blocked': []} for label, _ in targets.values()}
        for row in reversed(entries):
            module, snapshot = row['module'], row['snapshot']
            label, controller = targets[module]
            code = snapshot[0]
            try:
                current = self._existing(module, controller, code)
                if current is None:
                    result = 'absent'
                else:
                    if _record_fingerprint(current) != snapshot:
                        raise ErpError('ERP 资料在导入后已变更。')
                    self._request(f'baseInformation/{controller}/delete{controller}',
                                  params={'ModuleFunctionCode': module, 'data': code})
                    if self._existing(module, controller, code) is not None:
                        raise ErpError('ERP 仍可查到此色号。')
                    result = 'deleted'
            except ErpError as ex:
                results[label]['blocked'].append(f'{code}（{ex}）')
                if not self.logged_in:
                    break
                continue

            row['undo'] = result
            self._save_batch(batch)
            results[label][result].append(code)
        return results
