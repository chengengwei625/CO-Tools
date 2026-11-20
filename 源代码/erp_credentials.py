"""使用当前 Windows 用户的 DPAPI 加密保存 ERP 登录信息，不保存令牌。"""
import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
import tempfile


class CredentialError(Exception):
    pass


class _Blob(ctypes.Structure):
    _fields_ = [('size', wintypes.DWORD), ('data', ctypes.POINTER(ctypes.c_ubyte))]


def _crypt(value, protect):
    if os.name != 'nt':
        raise CredentialError('记住密码功能需要 Windows 用户加密服务。')
    buffer = ctypes.create_string_buffer(value)
    source = _Blob(len(value), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    target = _Blob()
    api = ctypes.WinDLL('crypt32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    if protect:
        call = api.CryptProtectData
        call.argtypes = [ctypes.POINTER(_Blob), wintypes.LPCWSTR, ctypes.c_void_p,
                         ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_Blob)]
        arguments = (ctypes.byref(source), 'CO ERP login', None, None, None, 1, ctypes.byref(target))
    else:
        call = api.CryptUnprotectData
        call.argtypes = [ctypes.POINTER(_Blob), ctypes.c_void_p, ctypes.c_void_p,
                         ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_Blob)]
        arguments = (ctypes.byref(source), None, None, None, None, 1, ctypes.byref(target))
    call.restype = wintypes.BOOL
    if not call(*arguments):
        raise CredentialError('无法读写已记住的登录信息，请在当前 Windows 用户下重新登录。')
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p
    try:
        return ctypes.string_at(target.data, target.size)
    finally:
        kernel.LocalFree(target.data)


class CredentialStore:
    def __init__(self, path=None):
        base = Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'COMerchandiserTool'
        self.path = Path(path) if path else base / 'erp_login.bin'

    def load(self):
        if not self.path.is_file():
            return None
        try:
            payload = self.path.read_bytes()
            if not payload.startswith(b'COERP1\0'):
                raise ValueError()
            result = json.loads(_crypt(payload[7:], False).decode('utf-8'))
            if not isinstance(result, dict) or not all(isinstance(result.get(k), str) and result[k]
                    for k in ('url', 'username', 'password', 'organization')):
                raise ValueError()
            return result
        except CredentialError:
            raise
        except Exception as ex:
            raise CredentialError('已保存的登录信息无法读取，请重新输入账号和密码。') from ex

    def save(self, url, username, password, organization='1'):
        if not username.strip() or not password:
            raise CredentialError('账号和密码不能为空。')
        value = {'url': url, 'username': username.strip(), 'password': password,
                 'organization': organization.strip() or '1'}
        protected = b'COERP1\0' + _crypt(json.dumps(value, ensure_ascii=False).encode('utf-8'), True)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='erp_login_', suffix='.tmp', dir=self.path.parent)
        try:
            with os.fdopen(fd, 'wb') as stream:
                stream.write(protected)
            os.replace(temporary, self.path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)

    def clear(self):
        self.path.unlink(missing_ok=True)
