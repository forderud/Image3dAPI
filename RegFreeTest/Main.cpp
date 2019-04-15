#include "../Image3dAPI/ComSupport.hpp"
#include "../Image3dAPI/IImage3d.h"
#include "../Image3dAPI/RegistryCheck.hpp"
#include <iostream>


void ParseSource (IImage3dSource & source) {
    Cart3dGeom geom = {};
    CHECK(source.GetBoundingBox(&geom));

    unsigned int frame_count = 0;
    CHECK(source.GetFrameCount(&frame_count));

    CComSafeArray<unsigned int> color_map;
    {
        SAFEARRAY * tmp = nullptr;
        CHECK(source.GetColorMap(&tmp));
        color_map.Attach(tmp);
        tmp = nullptr;
    }

    for (unsigned int frame = 0; frame < frame_count; ++frame) {
        unsigned short max_res[] = {64, 64, 64};

        // retrieve frame data
        Image3d data;
        CHECK(source.GetFrame(frame, geom, max_res, &data));
    }
}


int wmain (int argc, wchar_t *argv[]) {
    if (argc < 3) {
        std::wcout << L"Usage:\n";
        std::wcout << L"RegFreeTest.exe <loader-progid> <filename>" << std::endl;
        return -1;
    }

    CComBSTR progid = argv[1];  // e.g. "DummyLoader.Image3dFileLoader"
    CComBSTR filename = argv[2];

    ComInitialize com(COINIT_MULTITHREADED);

    CLSID clsid = {};
    CHECK(CLSIDFromProgID(progid, &clsid));

    // verify that loader library is compatible
    //CHECK(CheckImage3dAPIVersion(clsid)); // disabled, since it's not compatible with reg-free COM

    // create loader without accessing the registry (possible since the DummyLoader manifest is linked in)
    CComPtr<IImage3dFileLoader> loader;
    CHECK(loader.CoCreateInstance(clsid));

    {
        // load file
        Image3dError err_type = {};
        CComBSTR err_msg;
        CHECK(loader->LoadFile(filename, &err_type, &err_msg));
    }

    CComPtr<IImage3dSource> source;
    CHECK(loader->GetImageSource(&source));

    ProbeInfo probe;
    CHECK(source->GetProbeInfo(&probe));

    EcgSeries ecg;
    CHECK(source->GetECG(&ecg));

    ParseSource(*source);

    return 0;
}
