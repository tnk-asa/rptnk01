# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['C:\\Users\\P0244\\Documents\\py2exe\\20_pmdaXML2docx.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
#---------------’Ç‰Á‚±‚±‚©‚ç---------------
Key = ['mkl']

def remove_from_list(input, keys):
    outlist = []
    for item in input:
        name, _, _ = item
        flag = 0
        for key_word in keys:
            if name.find(key_word) > -1:
                flag = 1
        if flag != 1:
            outlist.append(item)
    return outlist

a.binaries = remove_from_list(a.binaries, Key)

#---------------’Ç‰Á‚±‚±‚Ü‚Å---------------

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='20_pmdaXML2docx',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
