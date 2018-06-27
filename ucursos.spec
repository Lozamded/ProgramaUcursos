# -*- mode: python -*-

block_cipher = None


a = Analysis(['ucursos.py'],
             pathex=['c:\\Users\\tomas\\Desktop\\programa_ucursos'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='ucursos',
          debug=False,
          strip=False,
          upx=True,
          console=True , icon='ind.ico')
