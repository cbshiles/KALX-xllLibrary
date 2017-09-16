# First cut of cmake builds

- No handling of the documentation etc is included.
- Only provides the logic to build the xll and test code using cmake under 32/64 bits.

# Install cmake

1. Download cmake, https://cmake.org/download/
2. Install cmake 

# Build Visual Studio 2015

Below the steps are outlined to build for 32 and 64 bit targets.
The builds below will put all build artefacts in the subdirectories `_bld32` and `_bld64`.

For example for the 32bit build, we have the following.
- The Visual Studio solution file `xll.sln` is located in `_bld32`.
- All build artefacts are in the directory `_bld32`.
  - The xll test addin is located in `_bld32/test/{Release,Debug}/`.
  - The xll core logic is located in `_bld32/xll/{Release,Debug}/`.

## Build 32

   ```
   "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\vcvarsall.bat" x86
   cmake -B_bld32 -H. -G "Visual Studio 14 2015"
   cmake --build _bld32 --config Release
   cmake --build _bld32 --config Debug
   ```

## Build 64

   ```
   "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\vcvarsall.bat" x64
   cmake -B_bld64 -H. -G "Visual Studio 14 2015 Win64"
   cmake --build _bld64 --config Release
   cmake --build _bld64 --config Debug
   ```


