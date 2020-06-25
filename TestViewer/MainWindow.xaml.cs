using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Image3dAPI;


/** Alternative to System.Activator.CreateInstance that allows explicit control over the activation context.
 *  Based on https://stackoverflow.com/questions/22901224/hosting-managed-code-and-garbage-collection */
public static class ComExt {
    [DllImport("ole32.dll", ExactSpelling = true, PreserveSig = false)]
    static extern void CoCreateInstance(
       [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
       [MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter,
       uint dwClsContext,
       [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
       [MarshalAs(UnmanagedType.Interface)] out object rReturnedComObject);

    public static object CreateInstance(Guid clsid, bool force_out_of_process) {
        Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
        const uint CLSCTX_INPROC_SERVER = 0x1;
        const uint CLSCTX_LOCAL_SERVER  = 0x4;

        uint class_context = CLSCTX_LOCAL_SERVER; // always allow out-of-process activation
        if (!force_out_of_process)
            class_context |= CLSCTX_INPROC_SERVER; // allow in-process activation

        object unk;
        CoCreateInstance(clsid, null, class_context, IID_IUnknown, out unk);
        return unk;
    }
}


namespace TestViewer
{
    public partial class MainWindow : Window
    {
        IImage3dFileLoader m_loader;
        IImage3dSource     m_source;

        Cart3dGeom         m_bboxXY;
        Cart3dGeom         m_bboxXZ;
        Cart3dGeom         m_bboxZY;

        public MainWindow()
        {
            InitializeComponent();
        }

        void ClearUI()
        {
            FrameSelector.Minimum = 0;
            FrameSelector.Maximum = 0;
            FrameSelector.IsEnabled = false;
            FrameSelector.Value = 0;

            FrameCount.Text = "";
            ProbeInfo.Text = "";
            InstanceUID.Text = "";


            ImageXY.Source = null;
            ImageXZ.Source = null;
            ImageZY.Source = null;

            m_bboxXY = new Cart3dGeom();
            m_bboxXZ = new Cart3dGeom();
            m_bboxZY = new Cart3dGeom();

            ECG.Data = null;

            if (m_source != null) {
                Marshal.ReleaseComObject(m_source);
                m_source = null;
            }
        }

        private void LoadDefaultBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadImpl(false);
        }

        private void LoadOutOfProcBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadImpl(true);
        }

        private void LoadImpl(bool force_out_of_proc)
        {
            // try to parse string as ProgId first
            Type comType = Type.GetTypeFromProgID(LoaderName.Text);
            if (comType == null) {
                try {
                    // fallback to parse string as CLSID hex value
                    Guid guid = Guid.Parse(LoaderName.Text);
                    comType = Type.GetTypeFromCLSID(guid);
                } catch (FormatException) {
                    MessageBox.Show("Unknown ProgId of CLSID.");
                    return;
                }
            }

            // API version compatibility check
            try {
                RegistryKey ver_key = Registry.ClassesRoot.OpenSubKey("CLSID\\{" + comType.GUID + "}\\Version");
                string ver_str = (string)ver_key.GetValue("");
                string cur_ver = string.Format("{0}.{1}", (int)Image3dAPIVersion.IMAGE3DAPI_VERSION_MAJOR, (int)Image3dAPIVersion.IMAGE3DAPI_VERSION_MINOR);
                if (ver_str != cur_ver) {
                    MessageBox.Show(string.Format("Loader uses version {0}, while the current version is {1}.", ver_str, cur_ver), "Incompatible loader version");
                    return;
                }
            } catch (Exception err) {
                MessageBox.Show(err.Message, "Version check error");
                // continue, since this error will also appear if the loader has non-matching bitness
            }

            // clear UI when switching to a new loader
            ClearUI();

            if (m_loader != null)
                Marshal.ReleaseComObject(m_loader);
            m_loader = (IImage3dFileLoader)ComExt.CreateInstance(comType.GUID, force_out_of_proc);

            this.FileOpenBtn.IsEnabled = true;
        }

        private void FileSelectBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() != true)
                return; // user hit cancel

            // clear UI when opening a new file
            ClearUI();

            FileName.Text = dialog.FileName;
        }

        private void FileOpenBtn_Click(object sender, RoutedEventArgs e)
        {
            Debug.Assert(m_loader != null);

            Image3dError err_type = Image3dError.Image3d_SUCCESS;
            string err_msg = "";
            try {
                m_loader.LoadFile(FileName.Text, out err_type, out err_msg);
            } catch (Exception) {
                // NOTE: err_msg does not seem to be marshaled back on LoadFile failure in .Net.
                // NOTE: This problem is limited to .Net, and does not occur in C++

                string message = "Unknown error";
                if ((err_type != Image3dError.Image3d_SUCCESS)) {
                    switch (err_type) {
                        case Image3dError.Image3d_ACCESS_FAILURE:
                            message = "Unable to open the file. The file might be missing or locked.";
                            break;
                        case Image3dError.Image3d_VALIDATION_FAILURE:
                            message = "Unsupported file. Probably due to unsupported vendor or modality.";
                            break;
                        case Image3dError.Image3d_NOT_YET_SUPPORTED:
                            message = "The loader is too old to parse the file.";
                            break;
                        case Image3dError.Image3d_SUPPORT_DISCONTINUED:
                            message = "The the file version is no longer supported (pre-DICOM format?).";
                            break;
                    }
                }
                MessageBox.Show("LoadFile error: " + message + " (" + err_msg+")");
                return;
            }

            try {
                if (m_source != null)
                    Marshal.ReleaseComObject(m_source);
                m_source = m_loader.GetImageSource();
            } catch (Exception err) {
                MessageBox.Show("ERROR: " + err.Message, "GetImageSource error");
                return;
            }

            FrameSelector.Minimum = 0;
            FrameSelector.Maximum = m_source.GetFrameCount()-1;
            FrameSelector.IsEnabled = true;
            FrameSelector.Value = 0;

            FrameCount.Text = "Frame count: " + m_source.GetFrameCount();
            ProbeInfo.Text = "Probe name: "+ m_source.GetProbeInfo().name;
            InstanceUID.Text = "UID: " + m_source.GetSopInstanceUID();

            InitializeSlices();
            DrawSlices(0);
            DrawEcg(m_source.GetFrameTimes()[0]);
        }

        private void DrawEcg (double cur_time)
        {
            EcgSeries ecg;
            try {
                ecg = m_source.GetECG();

                if (ecg.samples.Length == 0) {
                    ECG.Data = null; // ECG not available
                    return;
                }
            } catch (Exception) {
                ECG.Data = null; // ECG not available
                return;
            }

            // shrink width & height slightly, so that the "actual" width/height remain unchanged
            double W = (int)(ECG.ActualWidth - 1);
            double H = (int)(ECG.ActualHeight - 1);

            // horizontal scaling (index to X coord)
            double ecg_pitch = W/ecg.samples.Length;

            // vertical scaling (sample val to Y coord)
            double ecg_offset =  H*ecg.samples.Max()/(ecg.samples.Max()-ecg.samples.Min());
            double ecg_scale  = -H/(ecg.samples.Max()-ecg.samples.Min());

            // vertical scaling (time to Y coord conv)
            double time_offset = -W*ecg.start_time/(ecg.delta_time*ecg.samples.Length);
            double time_scale  =  W/(ecg.delta_time*ecg.samples.Length);

            PathGeometry pathGeom = new PathGeometry();
            {
                // draw ECG trace
                PathSegmentCollection pathSegmentCollection = new PathSegmentCollection();
                for (int i = 0; i < ecg.samples.Length; ++i) {
                    LineSegment lineSegment = new LineSegment();
                    lineSegment.Point = new Point(ecg_pitch*i, ecg_offset+ecg_scale*ecg.samples[i]);
                    pathSegmentCollection.Add(lineSegment);
                }

                PathFigure pathFig = new PathFigure();
                pathFig.StartPoint = new Point(0, ecg_offset+ecg_scale*ecg.samples[0]);
                pathFig.Segments = pathSegmentCollection;

                PathFigureCollection pathFigCol = new PathFigureCollection();
                pathFigCol.Add(pathFig);

                pathGeom.Figures = pathFigCol;
            }

            {
                // draw current frame line
                double x_pos = time_offset + time_scale * cur_time;

                LineGeometry line = new LineGeometry();
                line.StartPoint = new Point(x_pos, 0);
                line.EndPoint = new Point(x_pos, H);

                pathGeom.AddGeometry(line);
            }

            ECG.Stroke = Brushes.Blue;
            ECG.StrokeThickness = 1.0;
            ECG.Data = pathGeom;
        }

        private void FrameSelector_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var idx = (uint)FrameSelector.Value;
            DrawSlices(idx);
            DrawEcg(m_source.GetFrameTimes()[idx]);
        }

        private void InitializeSlices()
        {
            Debug.Assert(m_source != null);

            Cart3dGeom bbox = m_source.GetBoundingBox();
            if (Math.Abs(bbox.dir3[1]) > Math.Abs(bbox.dir2[1])) {
                // swap 2nd & 3rd axis, so that the 2nd becomes predominately "Y"
                SwapVals(ref bbox.dir2[0], ref bbox.dir3[0]);
                SwapVals(ref bbox.dir2[1], ref bbox.dir3[1]);
                SwapVals(ref bbox.dir2[2], ref bbox.dir3[2]);
            }

            // extend bounding-box axes, so that dir1, dir2 & dir3 have equal length
            ExtendBoundingBox(ref bbox);

            // get XY plane (assumes 1st axis is "X" and 2nd is "Y")
            m_bboxXY = CloneGeom(bbox);
            m_bboxXY.origin[0] = m_bboxXY.origin[0] + m_bboxXY.dir3[0] / 2;
            m_bboxXY.origin[1] = m_bboxXY.origin[1] + m_bboxXY.dir3[1] / 2;
            m_bboxXY.origin[2] = m_bboxXY.origin[2] + m_bboxXY.dir3[2] / 2;
            m_bboxXY.dir3[0] = 0;
            m_bboxXY.dir3[1] = 0;
            m_bboxXY.dir3[2] = 0;

            // get XZ plane (assumes 1st axis is "X" and 3rd is "Z")
            m_bboxXZ = CloneGeom(bbox);
            m_bboxXZ.origin[0] = m_bboxXZ.origin[0] + m_bboxXZ.dir2[0] / 2;
            m_bboxXZ.origin[1] = m_bboxXZ.origin[1] + m_bboxXZ.dir2[1] / 2;
            m_bboxXZ.origin[2] = m_bboxXZ.origin[2] + m_bboxXZ.dir2[2] / 2;
            m_bboxXZ.dir2[0] = m_bboxXZ.dir3[0];
            m_bboxXZ.dir2[1] = m_bboxXZ.dir3[1];
            m_bboxXZ.dir2[2] = m_bboxXZ.dir3[2];
            m_bboxXZ.dir3[0] = 0;
            m_bboxXZ.dir3[1] = 0;
            m_bboxXZ.dir3[2] = 0;

            // get ZY plane (assumes 2nd axis is "Y" and 3rd is "Z")
            m_bboxZY = CloneGeom(bbox);
            m_bboxZY.origin[0] = bbox.origin[0] + bbox.dir1[0] / 2;
            m_bboxZY.origin[1] = bbox.origin[1] + bbox.dir1[1] / 2;
            m_bboxZY.origin[2] = bbox.origin[2] + bbox.dir1[2] / 2;
            m_bboxZY.dir1[0] = bbox.dir3[0];
            m_bboxZY.dir1[1] = bbox.dir3[1];
            m_bboxZY.dir1[2] = bbox.dir3[2];
            m_bboxZY.dir2[0] = bbox.dir2[0];
            m_bboxZY.dir2[1] = bbox.dir2[1];
            m_bboxZY.dir2[2] = bbox.dir2[2];
            m_bboxZY.dir3[0] = 0;
            m_bboxZY.dir3[1] = 0;
            m_bboxZY.dir3[2] = 0;
        }

        private void DrawSlices (uint frame)
        {
            Debug.Assert(m_source != null);

            uint[] color_map = m_source.GetColorMap();

            // retrieve image slices
            const ushort HORIZONTAL_RES = 256;
            const ushort VERTICAL_RES = 256;

            // get XY plane (assumes 1st axis is "X" and 2nd is "Y")
            Image3d imageXY = m_source.GetFrame(frame, m_bboxXY, new ushort[] { HORIZONTAL_RES, VERTICAL_RES, 1 });
            ImageXY.Source = GenerateBitmap(imageXY, color_map);

            // get XZ plane (assumes 1st axis is "X" and 3rd is "Z")
            Image3d imageXZ = m_source.GetFrame(frame, m_bboxXZ, new ushort[] { HORIZONTAL_RES, VERTICAL_RES, 1 });
            ImageXZ.Source = GenerateBitmap(imageXZ, color_map);

            // get ZY plane (assumes 2nd axis is "Y" and 3rd is "Z")
            Image3d imageZY = m_source.GetFrame(frame, m_bboxZY, new ushort[] { HORIZONTAL_RES, VERTICAL_RES, 1 });
            ImageZY.Source = GenerateBitmap(imageZY, color_map);

            FrameTime.Text = "Frame time: " + imageXY.time;
        }

        private WriteableBitmap GenerateBitmap(Image3d t_img, uint[] t_map)
        {
            Debug.Assert(t_img.format == ImageFormat.FORMAT_U8);

            WriteableBitmap bitmap = new WriteableBitmap(t_img.dims[0], t_img.dims[1], 96.0, 96.0, PixelFormats.Rgb24, null);
            bitmap.Lock();
            unsafe {
                for (int y = 0; y < bitmap.Height; ++y) {
                    for (int x = 0; x < bitmap.Width; ++x) {
                        byte t_val = t_img.data[x + y * t_img.stride0];

                        // lookup tissue color
                        byte[] channels = BitConverter.GetBytes(t_map[t_val]);

                        // assign red, green & blue
                        byte* pixel = (byte*)bitmap.BackBuffer + x * (bitmap.Format.BitsPerPixel / 8) + y * bitmap.BackBufferStride;
                        pixel[0] = channels[0]; // red
                        pixel[1] = channels[1]; // green
                        pixel[2] = channels[2]; // blue
                        // discard alpha channel
                    }
                }
            }
            bitmap.AddDirtyRect(new Int32Rect(0, 0, bitmap.PixelWidth, bitmap.PixelHeight));
            bitmap.Unlock();
            return bitmap;
        }

        static void SwapVals(ref float v1, ref float v2)
        {
            float tmp = v1;
            v1 = v2;
            v2 = tmp;
        }

        static float VecLen(float x, float y, float z)
        {
            return (float)Math.Sqrt(x * x + y * y + z * z);
        }

        static float VecLen(Cart3dGeom g, int idx)
        {
            if (idx == 1)
                return VecLen(g.dir1[0], g.dir1[1], g.dir1[2]);
            else if (idx == 2)
                return VecLen(g.dir2[0], g.dir2[1], g.dir2[2]);
            else if (idx == 3)
                return VecLen(g.dir3[0], g.dir3[1], g.dir3[2]);

            throw new Exception("unsupported direction index");
        }

        /** Scale bounding-box, so that all axes share the same length.
         *  Also update the origin to keep the bounding-box centered. */
        static void ExtendBoundingBox(ref Cart3dGeom g)
        {
            float dir1_len = VecLen(g, 1);
            float dir2_len = VecLen(g, 2);
            float dir3_len = VecLen(g, 3);

            float max_len = Math.Max(dir1_len, Math.Max(dir2_len, dir3_len));

            if (dir1_len < max_len)
            {
                float delta = max_len - dir1_len;
                float dx, dy, dz;
                ScaleVector(g.dir1[0], g.dir1[1], g.dir1[2], delta, out dx, out dy, out dz);
                // scale up dir1 so that it becomes the same length as the other axes
                g.dir1[0] += dx;
                g.dir1[1] += dy;
                g.dir1[2] += dz;
                // move origin to keep the bounding-box centered
                g.origin[0] -= dx/2;
                g.origin[1] -= dy/2;
                g.origin[2] -= dz/2;
            }

            if (dir2_len < max_len)
            {
                float delta = max_len - dir2_len;
                float dx, dy, dz;
                ScaleVector(g.dir2[0], g.dir2[1], g.dir2[2], delta, out dx, out dy, out dz);
                // scale up dir2 so that it becomes the same length as the other axes
                g.dir2[0] += dx;
                g.dir2[1] += dy;
                g.dir2[2] += dz;
                // move origin to keep the bounding-box centered
                g.origin[0] -= dx / 2;
                g.origin[1] -= dy / 2;
                g.origin[2] -= dz / 2;
            }

            if (dir3_len < max_len)
            {
                float delta = max_len - dir3_len;
                float dx, dy, dz;
                ScaleVector(g.dir3[0], g.dir3[1], g.dir3[2], delta, out dx, out dy, out dz);
                // scale up dir3 so that it becomes the same length as the other axes
                float factor = max_len / dir3_len;
                g.dir3[0] += dx;
                g.dir3[1] += dy;
                g.dir3[2] += dz;
                // move origin to keep the bounding-box centered
                g.origin[0] -= dx / 2;
                g.origin[1] -= dy / 2;
                g.origin[2] -= dz / 2;
            }
        }

        static void ScaleVector(float in_x, float in_y, float in_z, float length, out float out_x, out float out_y, out float out_z)
        {
            float cur_len = VecLen(in_x, in_y, in_z);
            out_x = in_x * length / cur_len;
            out_y = in_y * length / cur_len;
            out_z = in_z * length / cur_len;
        }

        static Cart3dGeom CloneGeom (Cart3dGeom input)
        {
            Cart3dGeom copy;
            copy.origin = (float[])input.origin.Clone();
            copy.dir1 = (float[])input.dir1.Clone();
            copy.dir2 = (float[])input.dir2.Clone();
            copy.dir3 = (float[])input.dir3.Clone();
            return copy;
        }
    }
}
