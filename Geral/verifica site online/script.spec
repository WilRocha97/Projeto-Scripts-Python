# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['script.py'],
             pathex=['C:\\Users\\Rubens\\Desktop\\arquivos\\robos\\Gerais\\verifica site online'],
             binaries=[],
             datas=[('C:\\Users\\Rubens\\AppData\\Local\\Programs\\Python\\Python38-32\\lib\\site-packages\\eel\\eel.js', 'eel'), ('web', 'web')],
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
          name='Verifica Site',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='web\\favicon.ico')
