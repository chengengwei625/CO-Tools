"""发布版本：每段按十进制进位，1.1.9 的下一版为 1.2.0。"""

APP_NAME = 'CO 外贸跟单工具'
APP_VERSION = '1.1.3'
APP_TITLE = f'{APP_NAME}_{APP_VERSION}'


def next_version(version):
    major, minor, patch = map(int, version.split('.'))
    if major < 0 or not 0 <= minor <= 9 or not 0 <= patch <= 9:
        raise ValueError('版本格式须为非负整数主版本号及 0–9 的次版本号、修订号')
    minor, patch = divmod(minor * 10 + patch + 1, 10)
    carry, minor = divmod(minor, 10)
    return f'{major + carry}.{minor}.{patch}'
