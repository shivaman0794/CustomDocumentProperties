//****************************************************************************************************************************************//
// CopyRights       : Intergraph Corporation. All rights reserved
// Class            : CustomDocProperties
// Created By       : Aman Gupta
// Description      : This class can be used to Read, Add, delete and update custom properties(UserDefinderProperties) of the document.
// Date             : 09-23-2019 (MM-DD-YYYY)
//*****

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;


namespace CustomDocumentProperties
{
    
    public class CustomDocProperties : IDisposable
    {
        #region Class objects and GUID initialization
        //public static readonly Guid SummaryInformationFormatId = new Guid("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}");
        //public static readonly Guid DocSummaryInformationFormatId = new Guid("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}");
        public static readonly Guid UserDefinedPropertiesId = new Guid("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");

        private IList<StructuredProperty> _properties = new List<StructuredProperty>();
        public StructuredProperty this[string propertyName]
        {
            get
            {
               var property = _properties.SingleOrDefault(prop => prop.Name.Equals(propertyName));
                if (property == null)
                {
                    return null;
                }
                return property;
            }
        }
        public string FilePath { get; private set; }

        private int propCount = 0;
        public IList<StructuredProperty> Properties
        {
            get
            {
                return _properties;
            }
        }

        #endregion

        #region CustomDocProperties class cunstructor

        /// <summary>
        /// Class constructor which loads all custom properties.
        /// </summary>
        /// <param name="filePath">Path of the document</param>
        public CustomDocProperties(string filePath)
        {
            if (filePath == null)
                throw new ArgumentNullException("filePath");

            FilePath = filePath;

            IPropertySetStorage propertySetStorage = null;
            IPropertyStorage propertyStorage = null;

            try
            {
                int hr = StgOpenStorageEx(FilePath, STGM.STGM_READ | STGM.STGM_SHARE_DENY_NONE | STGM.STGM_DIRECT_SWMR, STGFMT.STGFMT_ANY, 0, IntPtr.Zero, IntPtr.Zero, typeof(IPropertySetStorage).GUID, out propertySetStorage);
                if (hr == STG_E_FILENOTFOUND || hr == STG_E_PATHNOTFOUND)
                    throw new FileNotFoundException(null, FilePath);

                if (hr != 0)
                    throw new Win32Exception(hr);

                LoadPropertySet(propertySetStorage, UserDefinedPropertiesId);
                
            }
            finally
            {
                if (propertySetStorage != null)
                {
                    Marshal.ReleaseComObject(propertySetStorage);
                }
                if (propertyStorage != null)
                {
                    Marshal.ReleaseComObject(propertyStorage);
                }
            }
            
        }

        #endregion

        #region All methods to modify the custom properties

        /// <summary>
        /// Adds a new Custom Property to the document.
        /// </summary>
        /// <param name="name">Property Name to be added</param>
        /// <param name="value">Property Value to be added</param>
        public void AddProperty(string name, object value)
        {
            IPropertySetStorage propertySetStorage = null;
            IPropertyStorage propertyStorage = null;
            try
            {
                bool doesPropertyExist = _properties.Any(x => x.Name == name);
                if (doesPropertyExist)
                {
                    throw new Win32Exception("Property named  \'"+name+ "\' already exists");
                }

                //The StgOpenStorageEx function opens an existing root storage object in the file system.
                int hr = StgOpenStorageEx(FilePath, STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_READWRITE, STGFMT.STGFMT_ANY, 0, IntPtr.Zero, IntPtr.Zero, typeof(IPropertySetStorage).GUID, out propertySetStorage);
                if (hr == STG_E_FILENOTFOUND || hr == STG_E_PATHNOTFOUND)
                    throw new FileNotFoundException(null, FilePath);



                if (hr != 0)
                    throw new Win32Exception(hr);

                //The Create method creates and opens a new property set in the property set storage object
                hr = propertySetStorage.Create(UserDefinedPropertiesId, UserDefinedPropertiesId, 0, STGM.STGM_CREATE | STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE, out propertyStorage);
                if (hr == STG_E_ACCESSDENIED)
                    return;

                if (hr != 0)
                    throw new Win32Exception(hr);

                //Specify a property either by its property identifier (ID) or the associated string name.
                var propspec = new PROPSPEC[1];
                propspec[0] = new PROPSPEC();
                propspec[0].ulKind = PRSPEC.PRSPEC_LPWSTR;
                propspec[0].union.lpwstr = Marshal.StringToHGlobalUni(name.ToString());

                //Define the type tag and the value of a property in a property set.
                var vars = new PROPVARIANT[1];
                vars[0] = new PROPVARIANT();
                vars[0].vt = VARTYPE.VT_LPWSTR;
                vars[0].union.pwszVal = Marshal.StringToHGlobalUni(value.ToString());

                /// <summary>
                /// The WriteMultiple method writes a specified group of properties to the current property set.
                /// </summary>
                /// <param name="cpspec=1">The number of properties set</param>
                /// <param name="propspec"></param>
                /// <param name="vars"></param>
                /// <param name="propspec =2">The minimum value for the property IDs that the method must assign</param>

                hr = propertyStorage.WriteMultiple(1, propspec, vars, 2);

                if (hr == 0)
                {
                    hr = propertyStorage.Commit(0);
                }
                

                if (hr == 0)
                {
                    var property = new StructuredProperty(UserDefinedPropertiesId, name, propCount++);
                    property.Value = value;
                    _properties.Add(property);
                }

            }

            finally
            {
                if (propertySetStorage != null)
                {
                    Marshal.ReleaseComObject(propertySetStorage);
                }
                if (propertyStorage != null)
                {
                    Marshal.ReleaseComObject(propertyStorage);
                }
            }
        }

        /// <summary>
        /// Updates the existing property with the passed value to the document.
        /// </summary>
        /// <param name="name">Property Name to be updated</param>
        /// <param name="value">Property Value to be Updated</param>
        public void UpdateProperty(string name, object value)
        {
            IPropertySetStorage propertySetStorage = null;
            IPropertyStorage propertyStorage = null;
            try
            {
                bool doesPropertyExist = _properties.Any(x => x.Name == name);

                if (doesPropertyExist)
                {
                    
                    int hr = StgOpenStorageEx(FilePath, STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_READWRITE, STGFMT.STGFMT_ANY, 0, IntPtr.Zero, IntPtr.Zero, typeof(IPropertySetStorage).GUID, out propertySetStorage);
                    if (hr == STG_E_FILENOTFOUND || hr == STG_E_PATHNOTFOUND)
                        throw new FileNotFoundException(null, FilePath);

                    if (hr != 0)
                        throw new Win32Exception(hr);

                    hr = propertySetStorage.Create(UserDefinedPropertiesId, UserDefinedPropertiesId, 0, STGM.STGM_CREATE | STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE, out propertyStorage);
                    if (hr == STG_E_ACCESSDENIED)
                        return;

                    if (hr != 0)
                        throw new Win32Exception(hr);

                    var propspec = new PROPSPEC[1];
                    propspec[0] = new PROPSPEC();
                    propspec[0].ulKind = PRSPEC.PRSPEC_LPWSTR;
                    propspec[0].union.lpwstr = Marshal.StringToHGlobalUni(name.ToString());

                    var vars = new PROPVARIANT[1];
                    vars[0] = new PROPVARIANT();
                    vars[0].vt = VARTYPE.VT_LPWSTR;
                    vars[0].union.pwszVal = Marshal.StringToHGlobalUni(value.ToString());
                    hr = propertyStorage.WriteMultiple(1, propspec, vars, 2);

                    if (hr == 0)
                    {
                        hr = propertyStorage.Commit(0);
                        if (hr == 0)
                        {
                            var prop = _properties.Single(x => x.Name == name);
                            prop.Value = value;
                        }
                    }
                }

                else
                {
                    throw new Exception("Document does not contain any custom property named \""+name+"\'");
                }
            }

            finally
            {
                if (propertySetStorage != null)
                {
                    Marshal.ReleaseComObject(propertySetStorage);
                }
                if (propertyStorage != null)
                {
                    Marshal.ReleaseComObject(propertyStorage);
                }
            }
        }

        /// <summary>
        /// Deletes Custom Property if it exists.
        /// </summary>
        /// <param name="name">Property to be deleted from the document using property name</param>
        public void DeleteProperty(string name)
        {
            IPropertySetStorage propertySetStorage = null;
            IPropertyStorage propertyStorage = null;

            try
            {
                bool doesPropertyExist = _properties.Any(x => x.Name == name);
                if (!doesPropertyExist)
                {
                    throw new Win32Exception("Document does not contain property named  \'" + name + "\' ");
                }

                int hr = StgOpenStorageEx(FilePath, STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_READWRITE, STGFMT.STGFMT_ANY, 0, IntPtr.Zero, IntPtr.Zero, typeof(IPropertySetStorage).GUID, out propertySetStorage);
                if (hr == STG_E_FILENOTFOUND || hr == STG_E_PATHNOTFOUND)
                    throw new FileNotFoundException(null, FilePath);

                if (hr != 0)
                    throw new Win32Exception(hr);

                hr = propertySetStorage.Create(UserDefinedPropertiesId, UserDefinedPropertiesId, 0, STGM.STGM_CREATE | STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE, out propertyStorage);
                if (hr == STG_E_ACCESSDENIED)
                    return;

                if (hr != 0)
                    throw new Win32Exception(hr);

                var propspec = new PROPSPEC[1];
                propspec[0] = new PROPSPEC();
                propspec[0].ulKind = PRSPEC.PRSPEC_LPWSTR;
                propspec[0].union.lpwstr = Marshal.StringToHGlobalUni(name.ToString());

                hr = propertyStorage.DeleteMultiple(1, propspec);

                if (hr == 0)
                {
                    hr = propertyStorage.Commit(0);
                }

                if (hr == 0)
                {
                    propCount--;
                    _properties.Remove(_properties.FirstOrDefault(x => x.Name == name));
                }
            }

            finally
            {
                if (propertySetStorage != null)
                {
                    Marshal.ReleaseComObject(propertySetStorage);
                }
                if (propertyStorage != null)
                {
                    Marshal.ReleaseComObject(propertyStorage);
                }
            }
        }

        /// <summary>
        /// Opens the document and reads user defined properties
        /// </summary>
        /// <param name="propertySetStorage">Out parameter of StgOpenStorageEx method</param>
        /// <param name="fmtid">UserDefinedPropertiesId</param>
        private void LoadPropertySet(IPropertySetStorage propertySetStorage, Guid fmtid)
        {
            IPropertyStorage propertyStorage;
            Guid guid = new Guid("{14A30E07-3193-4083-8C03-38ABA4A316A8}");
            
            int hr = propertySetStorage.Open(fmtid, (STGM.STGM_READ | STGM.STGM_SHARE_EXCLUSIVE),out propertyStorage);
            if (hr == STG_E_FILENOTFOUND || hr == STG_E_ACCESSDENIED)
                return;

            if (hr != 0)
                throw new Win32Exception(hr);

            IEnumSTATPROPSTG es;
            propertyStorage.Enum(out es);
            if (es == null)
                return;

            try
            {
                var stg = new STATPROPSTG();
                int fetched;
                do
                {
                    hr = es.Next(1, ref stg, out fetched);
                    if (hr != 0 && hr != 1)
                        throw new Win32Exception(hr);

                    if (fetched == 1)
                    {
                        string name = GetPropertyName(fmtid, propertyStorage, stg);

                        var propsec = new PROPSPEC[1];
                        propsec[0] = new PROPSPEC();
                        propsec[0].ulKind = stg.lpwstrName != null ? PRSPEC.PRSPEC_LPWSTR : PRSPEC.PRSPEC_PROPID;
                        IntPtr lpwstr = IntPtr.Zero;
                        if (stg.lpwstrName != null)
                        {
                            lpwstr = Marshal.StringToCoTaskMemUni(stg.lpwstrName);
                            propsec[0].union.lpwstr = lpwstr;
                        }
                        else
                        {
                            propsec[0].union.propid = stg.propid;
                        }

                        var vars = new PROPVARIANT[1];
                        vars[0] = new PROPVARIANT();
                        try
                        {
                            /// <summary>
                            /// The ReadMultiple method reads specified properties from the current property set.
                            /// </summary>
                            /// <param name="cpspec=1">The number of properties set</param>
                            /// <param name="propspec"></param>
                            /// <param name="vars"></param>

                            hr = propertyStorage.ReadMultiple(1, propsec, vars);
                            if (hr != 0)
                                throw new Win32Exception(hr);
                        }
                        finally
                        {
                            if (lpwstr != IntPtr.Zero)
                            {
                                Marshal.FreeCoTaskMem(lpwstr);
                            }
                        }

                        object value;
                        try
                        {
                            switch (vars[0].vt)
                            {
                                case VARTYPE.VT_BOOL:
                                    value = vars[0].union.boolVal != 0 ? true : false;
                                    break;

                                case VARTYPE.VT_BSTR:
                                    value = Marshal.PtrToStringUni(vars[0].union.bstrVal);
                                    break;

                                case VARTYPE.VT_CY:
                                    value = decimal.FromOACurrency(vars[0].union.cyVal);
                                    break;

                                case VARTYPE.VT_DATE:
                                    value = DateTime.FromOADate(vars[0].union.date);
                                    break;

                                case VARTYPE.VT_DECIMAL:
                                    IntPtr dec = IntPtr.Zero;
                                    Marshal.StructureToPtr(vars[0], dec, false);
                                    value = Marshal.PtrToStructure(dec, typeof(decimal));
                                    break;

                                case VARTYPE.VT_DISPATCH:
                                    value = Marshal.GetObjectForIUnknown(vars[0].union.pdispVal);
                                    break;

                                case VARTYPE.VT_ERROR:
                                case VARTYPE.VT_HRESULT:
                                    value = vars[0].union.scode;
                                    break;

                                case VARTYPE.VT_FILETIME:
                                    value = DateTime.FromFileTime(vars[0].union.filetime);
                                    break;

                                case VARTYPE.VT_I1:
                                    value = vars[0].union.cVal;
                                    break;

                                case VARTYPE.VT_I2:
                                    value = vars[0].union.iVal;
                                    break;

                                case VARTYPE.VT_I4:
                                    value = vars[0].union.lVal;
                                    break;

                                case VARTYPE.VT_I8:
                                    value = vars[0].union.hVal;
                                    break;

                                case VARTYPE.VT_INT:
                                    value = vars[0].union.intVal;
                                    break;

                                case VARTYPE.VT_LPSTR:
                                    value = Marshal.PtrToStringAnsi(vars[0].union.pszVal);
                                    break;

                                case VARTYPE.VT_LPWSTR:
                                    value = Marshal.PtrToStringUni(vars[0].union.pwszVal);
                                    break;

                                case VARTYPE.VT_R4:
                                    value = vars[0].union.fltVal;
                                    break;

                                case VARTYPE.VT_R8:
                                    value = vars[0].union.dblVal;
                                    break;

                                case VARTYPE.VT_UI1:
                                    value = vars[0].union.bVal;
                                    break;

                                case VARTYPE.VT_UI2:
                                    value = vars[0].union.uiVal;
                                    break;

                                case VARTYPE.VT_UI4:
                                    value = vars[0].union.ulVal;
                                    break;

                                case VARTYPE.VT_UI8:
                                    value = vars[0].union.uhVal;
                                    break;

                                case VARTYPE.VT_UINT:
                                    value = vars[0].union.uintVal;
                                    break;

                                case VARTYPE.VT_UNKNOWN:
                                    value = Marshal.GetObjectForIUnknown(vars[0].union.punkVal);
                                    break;

                                default:
                                    value = null;
                                    break;
                            }
                        }
                        finally
                        {
                            PropVariantClear(ref vars[0]);
                        }

                        var property = new StructuredProperty(fmtid, name, stg.propid);
                        property.Value = value;
                        _properties.Add(property);
                    }
                }
                while (fetched == 1);
                propCount = _properties.Count;
            }
            finally
            {
                Marshal.ReleaseComObject(es);
                Marshal.ReleaseComObject(propertyStorage);
            }
        }

        private static string GetPropertyName(Guid fmtid, IPropertyStorage propertyStorage, STATPROPSTG stg)
        {
            if (!string.IsNullOrEmpty(stg.lpwstrName))
                return stg.lpwstrName;

            var propids = new int[1];
            propids[0] = stg.propid;
            var names = new string[1];
            names[0] = null;
            int hr = propertyStorage.ReadPropertyNames(1, propids, names);
            if (hr == 0)
                return names[0];

            return null;
        }

        #endregion

        #region Variables, enums and structs

        public const int STG_E_FILENOTFOUND = unchecked((int)0x80030002);
        public const int STG_E_PATHNOTFOUND = unchecked((int)0x80030003);
        public const int STG_E_ACCESSDENIED = unchecked((int)0x80030005);
        public enum PRSPEC
        {
            PRSPEC_LPWSTR = 0,
            PRSPEC_PROPID = 1
        }
        public enum STGFMT
        {
            STGFMT_ANY = 4,
        }

        [Flags]
        public enum STGM
        {
            STGM_READ = 0x00000000,
            STGM_READWRITE = 0x00000002,
            STGM_SHARE_DENY_NONE = 0x00000040,
            STGM_SHARE_DENY_WRITE = 0x00000020,
            STGM_SHARE_EXCLUSIVE = 0x00000010,
            STGM_DIRECT_SWMR = 0x00400000,
            STGM_CREATE = 0x00001000,
        }
        public enum PROPSETFLAG : int
        {
            DEFAULT = 0,
            NONSIMPLE = 1,
            ANSI = 2,
            UNBUFFERED = 4,
            CASE_SENSITIVE = 8
        }
        // we only define what we handle
        public enum VARTYPE : short
        {
            VT_I2 = 2,
            VT_I4 = 3,
            VT_R4 = 4,
            VT_R8 = 5,
            VT_CY = 6,
            VT_DATE = 7,
            VT_BSTR = 8,
            VT_DISPATCH = 9,
            VT_ERROR = 10,
            VT_BOOL = 11,
            VT_UNKNOWN = 13,
            VT_DECIMAL = 14,
            VT_I1 = 16,
            VT_UI1 = 17,
            VT_UI2 = 18,
            VT_UI4 = 19,
            VT_I8 = 20,
            VT_UI8 = 21,
            VT_INT = 22,
            VT_UINT = 23,
            VT_HRESULT = 25,
            VT_LPSTR = 30,
            VT_LPWSTR = 31,
            VT_FILETIME = 64,
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct PROPVARIANTunion
        {
            [FieldOffset(0)]
            public sbyte cVal;
            [FieldOffset(0)]
            public byte bVal;
            [FieldOffset(0)]
            public short iVal;
            [FieldOffset(0)]
            public ushort uiVal;
            [FieldOffset(0)]
            public int lVal;
            [FieldOffset(0)]
            public uint ulVal;
            [FieldOffset(0)]
            public int intVal;
            [FieldOffset(0)]
            public uint uintVal;
            [FieldOffset(0)]
            public long hVal;
            [FieldOffset(0)]
            public float fltVal;
            [FieldOffset(0)]
            public ulong uhVal;
            [FieldOffset(0)]
            public double dblVal;
            [FieldOffset(0)]
            public short boolVal;
            [FieldOffset(0)]
            public int scode;
            [FieldOffset(0)]
            public long cyVal;
            [FieldOffset(0)]
            public double date;
            [FieldOffset(0)]
            public long filetime;
            [FieldOffset(0)]
            public IntPtr bstrVal;
            [FieldOffset(0)]
            public IntPtr pszVal;
            [FieldOffset(0)]
            public IntPtr pwszVal;
            [FieldOffset(0)]
            public IntPtr punkVal;
            [FieldOffset(0)]
            public IntPtr pdispVal;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROPSPEC
        {
            public PRSPEC ulKind;
            public PROPSPECunion union;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct PROPSPECunion
        {
            [FieldOffset(0)]
            public int propid;
            [FieldOffset(0)]
            public IntPtr lpwstr;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROPVARIANT
        {
            public VARTYPE vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public PROPVARIANTunion union;
        }

        [StructLayout(LayoutKind.Explicit, Size = 16)]
        public struct PropVariant
        {
            [FieldOffset(0)] public short variantType;
            [FieldOffset(8)] public IntPtr pointerValue;
            [FieldOffset(8)] public byte byteValue;
            [FieldOffset(8)] public long longValue;

            public void FromObject(object obj)
            {
                if (obj.GetType() == typeof(string))
                {
                    this.variantType = (short)VarEnum.VT_LPWSTR;
                    this.pointerValue = Marshal.StringToHGlobalUni((string)obj);
                }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct STATPROPSTG
        {
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpwstrName;
            public int propid;
            public VARTYPE vt;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct STATPROPSETSTG
        {
            public Guid fmtid;
            public Guid clsid;
            public uint grfFlags;
            public System.Runtime.InteropServices.ComTypes.FILETIME mtime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ctime;
            public System.Runtime.InteropServices.ComTypes.FILETIME atime;
            public uint dwOSVersion;
        }

        #endregion

        #region System ole32.dll and storage references

        [DllImport("ole32.dll")]
        public static extern int StgOpenStorageEx([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, STGM grfMode, STGFMT stgfmt, int grfAttrs, IntPtr pStgOptions, IntPtr reserved2, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IPropertySetStorage ppObjectOpen);

        [DllImport("ole32.dll")]
        public static extern int PropVariantClear(ref PROPVARIANT pvar);

        [Guid("0000013B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumSTATPROPSETSTG
        {
            [PreserveSig]
            int Next(int celt, ref STATPROPSETSTG rgelt, out int pceltFetched);
            // rest ommited
        }

        [Guid("00000139-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumSTATPROPSTG
        {
            [PreserveSig]
            int Next(int celt, ref STATPROPSTG rgelt, out int pceltFetched);
            // rest ommited
        }

        [Guid("00000138-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStorage
        {
            [PreserveSig]
            int ReadMultiple(uint cpspec, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar);
            [PreserveSig]
            int WriteMultiple(uint cpspec, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]  PROPSPEC[] rgpspec, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]  PROPVARIANT[] rgpropvar, uint propidNameFirst);
            [PreserveSig]
            int DeleteMultiple(uint cpspec, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec);
            [PreserveSig]
            int ReadPropertyNames(uint cpropid, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] int[] rgpropid, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 0)] string[] rglpwstrName);
            [PreserveSig]
            int NotDeclared1();
            [PreserveSig]
            int NotDeclared2();
            [PreserveSig]
            int Commit(uint grfCommitFlags);
            [PreserveSig]
            int NotDeclared3();
            [PreserveSig]
            int Enum(out IEnumSTATPROPSTG ppenum);
            // rest ommited
        }

        [Guid("0000013A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertySetStorage
        {
            [PreserveSig]
            int Create([MarshalAs(UnmanagedType.LPStruct)] Guid rfmtid, [MarshalAs(UnmanagedType.LPStruct)] Guid pclsid, uint grfFlags, STGM grfMode, out IPropertyStorage ppprstg);
            [PreserveSig]
            int Open([MarshalAs(UnmanagedType.LPStruct)] Guid rfmtid, STGM grfMode, out IPropertyStorage ppprstg);
            [PreserveSig]
            int NotDeclared3();
            [PreserveSig]
            int Enum(out IEnumSTATPROPSETSTG ppenum);
        }

        #endregion

        public void Dispose()
        {
            FilePath = string.Empty;
        }
    }

    #region Structured proeprty class
    public class StructuredProperty : IDisposable
    {
        public StructuredProperty(Guid formatId, string name, int id)
        {
            FormatId = formatId;
            Name = name;
            Id = id;
        }

        public Guid FormatId { get; private set; }
        public string Name { get; private set; }
        public int Id { get; private set; }
        public object Value { get; set; }

        public void Dispose()
        {
            Value = null;
            Name = string.Empty;
        }
    }

    #endregion
}


