## Sample code to demonstrate how to access Image3dAPI from a python script
import platform
import comtypes
import comtypes.client
import numpy as np
from utils import SafeArrayToNumpy
from utils import FrameTo3dArray
from utils import TypeLibFromObject


if __name__=="__main__":
    # create loader object
    loader = comtypes.client.CreateObject("DummyLoader.Image3dFileLoader")
    # cast to IImage3dFileLoader interface
    Image3dAPI = TypeLibFromObject(loader)
    loader = loader.QueryInterface(Image3dAPI.IImage3dFileLoader)

    # load file
    err_type, err_msg = loader.LoadFile("dummy.dcm")
    source = loader.GetImageSource()

    # retrieve probe info
    probe = source.GetProbeInfo()
    print("Probe name: "+probe.name)
    print("Probe type: "+str(probe.type))

    # retrieve ECG info
    ecg = source.GetECG()
    print("ECG start: "+str(ecg.start_time))
    print("ECG delta: "+str(ecg.delta_time))
    samples = SafeArrayToNumpy(ecg.samples)
    trig_times = SafeArrayToNumpy(ecg.trig_times)

    # get bounding box
    bbox = source.GetBoundingBox()
    origin = [bbox.origin[0], bbox.origin[1], bbox.origin[2]]
    dir1   = [bbox.dir1[0],   bbox.dir1[1],   bbox.dir1[2]]
    dir2   = [bbox.dir2[0],   bbox.dir2[1],   bbox.dir2[2]]
    dir3   = [bbox.dir3[0],   bbox.dir3[1],   bbox.dir3[2]]

    color_map = source.GetColorMap()
    print("Color-map length: "+str(len(color_map)))

    frame_count = source.GetFrameCount()
    for i in range(frame_count):
        max_res = np.ctypeslib.as_ctypes(np.array([64, 64, 64], dtype=np.ushort))
        frame = source.GetFrame(i, bbox, max_res)
        
        print("Frame metadata:")
        print("  time:   "+str(frame.time))
        print("  format: "+str(frame.format))
        print("  dims:  ("+str(frame.dims[0])+", "+str(frame.dims[1])+", "+str(frame.dims[2])+")")

        data = FrameTo3dArray(frame)
        print("  shape: "+str(data.shape))

