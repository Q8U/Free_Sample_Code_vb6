
DLS Mediaps.dll: dlldata.obj DLS Media_p.obj DLS Media_i.obj
	link /dll /out:DLS Mediaps.dll /def:DLS Mediaps.def /entry:DllMain dlldata.obj DLS Media_p.obj DLS Media_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del DLS Mediaps.dll
	@del DLS Mediaps.lib
	@del DLS Mediaps.exp
	@del dlldata.obj
	@del DLS Media_p.obj
	@del DLS Media_i.obj
