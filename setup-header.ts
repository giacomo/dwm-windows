import { createRequire } from 'module';
import fs from 'fs';
import path from 'path';

const require = createRequire(import.meta.url);

try {
  const napiPath = require.resolve('node-addon-api');
  const napiDir = path.dirname(napiPath);
  const includePath = path.join(napiDir, 'napi.h');
  
  console.log('node-addon-api directory:', napiDir);
  console.log('napi.h exists:', fs.existsSync(includePath));
  
  if (fs.existsSync(includePath)) {
    console.log('✅ Headers found at:', napiDir);
  } else {
    console.log('❌ Headers not found');
    const files = fs.readdirSync(napiDir);
    console.log('Available files:', files.filter(f => f.endsWith('.h')));
    console.log('All files:', files);
  }
  
  // Get the correct path with forward slashes for gyp
  const correctPath = napiDir.replace(/\\/g, '/');
  
  // Create binding.gyp with the correct absolute path
  const bindingGyp = {
    targets: [{
      target_name: "dwm_windows",
      sources: ["dwm_thumbnail.cc"],
      include_dirs: [
        correctPath,
        "./node_modules/node-addon-api"
      ],
      libraries: [
        "dwmapi.lib",
        "psapi.lib", 
        "gdiplus.lib"
      ],
      defines: [
        "NAPI_DISABLE_CPP_EXCEPTIONS"
      ],
      "cflags!": [ "-fno-exceptions" ],
      "cflags_cc!": [ "-fno-exceptions" ],
      msvs_settings: {
        VCCLCompilerTool: {
          ExceptionHandling: 1
        }
      }
    }]
  };
  
  fs.writeFileSync('binding.gyp', JSON.stringify(bindingGyp, null, 2));
  console.log('✅ Updated binding.gyp with path:', correctPath);
  
} catch (error) {
  console.error('❌ Error:', (error as Error).message);
}