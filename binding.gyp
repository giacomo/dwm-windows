{
    "targets": [
        {
            "target_name": "dwm_windows",
            "sources": ["dwm_thumbnail.cc"],
            "include_dirs": [
                "C:/Projects/dwm-windows/node_modules/node-addon-api",
                "./node_modules/node-addon-api",
                "<(module_root_dir)/node_modules/node-addon-api",
            ],
            "libraries": [
                "dwmapi.lib",
                "psapi.lib",
                "gdiplus.lib",
                "shell32.lib",
                "propsys.lib",
            ],
            "defines": [
                "NAPI_DISABLE_CPP_EXCEPTIONS",
                "WIN32_LEAN_AND_MEAN",
                "NOMINMAX",
                "_WIN32_WINNT=0x0601",
            ],
            "cflags!": ["-fno-exceptions"],
            "cflags_cc!": ["-fno-exceptions"],
            "msvs_settings": {
                "VCCLCompilerTool": {"ExceptionHandling": 1, "AdditionalOptions": []}
            },
        }
    ]
}
