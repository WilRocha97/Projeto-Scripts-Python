# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files=[
    ('web', 'web'),
    ('eel', 'eel'),
]

a = Analysis(['script.py'],
             pathex=['.'],
             binaries=[],
             datas=added_files,
             hiddenimports=['bottle_websocket'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='separador arquivos dp',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='./web/icon.ico')
