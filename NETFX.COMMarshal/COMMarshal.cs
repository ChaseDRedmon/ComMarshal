using System;
using System.Runtime.InteropServices;
using RGiesecke.DllExport;

namespace COMMarshal
{
    [ComVisible(true)]
    [Guid("0e4d4d37-4edd-4ac6-80ce-ca7bc67a4238")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICOMMarshal
    {
        void MethodA();

        [return: MarshalAs(UnmanagedType.Bool)]
        bool MethodB();

        [return: MarshalAs(UnmanagedType.I4)]
        int MethodC();

        void MethodD([Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] out int[] a);
        void MethodE([Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] out string[] a);
        
        string PropertyA 
        { 
            [return: MarshalAs(UnmanagedType.BStr)] 
            get;

            [param: MarshalAs(UnmanagedType.BStr)]
            set; 
        }
        
        int PropertyB 
        {
            [return: MarshalAs(UnmanagedType.I4)]
            get;

            [param: MarshalAs(UnmanagedType.I4)]
            set; 
        }
        
        bool PropertyC 
        {
            [return: MarshalAs(UnmanagedType.Bool)]
            get;

            [param: MarshalAs(UnmanagedType.Bool)]
            set;
        }

        void ArgumentA([In, MarshalAs(UnmanagedType.I4)] int A);
        void ArgumentB([In, MarshalAs(UnmanagedType.R4)] float B);
        void ArgumentC([In, MarshalAs(UnmanagedType.R8)] double C);
        void ArgumentD([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] D);
        void ArgumentE([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] ref int[] E);
        
        void MethodInOutA([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] D);
        void MethodInOutB([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] ref int[] E);
        
        [return: MarshalAs(UnmanagedType.I4)]
        int MethodOutA([In, MarshalAs(UnmanagedType.I4)] int A, [Out, MarshalAs(UnmanagedType.I4)] out int B);
    }

    [ComVisible(true)]
    [Guid("a3ea8ead-cc9e-432b-a4e3-6810a2dee56e")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class COMMarshal : ICOMMarshal
    {
        public string PropertyA 
        {
            [return: MarshalAs(UnmanagedType.BStr)]
            get => throw new NotImplementedException();

            [param: MarshalAs(UnmanagedType.BStr)]
            set => throw new NotImplementedException();
        }

        public int PropertyB 
        {
            [return: MarshalAs(UnmanagedType.I4)]
            get => throw new NotImplementedException();

            [param: MarshalAs(UnmanagedType.I4)]
            set => throw new NotImplementedException();
        }

        public bool PropertyC 
        {
            [return: MarshalAs(UnmanagedType.Bool)]
            get => throw new NotImplementedException();

            [param: MarshalAs(UnmanagedType.Bool)]
            set => throw new NotImplementedException();
        }

        public void ArgumentA([In, MarshalAs(UnmanagedType.I4)] int A)
        {
            throw new NotImplementedException();
        }

        public void ArgumentB([In, MarshalAs(UnmanagedType.R4)] float B)
        {
            throw new NotImplementedException();
        }

        public void ArgumentC([In, MarshalAs(UnmanagedType.R8)] double C)
        {
            throw new NotImplementedException();
        }

        public void ArgumentD([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] D)
        {
            throw new NotImplementedException();
        }

        public void ArgumentE([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] ref int[] E)
        {
            throw new NotImplementedException();
        }

        public void MethodInOutA([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] D)
        {
            throw new NotImplementedException();
        }

        public void MethodInOutB([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] ref int[] E)
        {
            throw new NotImplementedException();
        }

        [return: MarshalAs(UnmanagedType.I4)]
        public int MethodOutA([In, MarshalAs(UnmanagedType.I4)] int A, [Out, MarshalAs(UnmanagedType.I4)] out int B)
        {
            B = A * 2;
            return B * B;
        }

        public void MethodA()
        {
            throw new NotImplementedException();
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        public bool MethodB()
        {
            throw new NotImplementedException();
        }

        [return: MarshalAs(UnmanagedType.I4)]
        public int MethodC()
        {
            throw new NotImplementedException();
        }

        public void MethodD([Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] out int[] a)
        {
            throw new NotImplementedException();
        }

        public void MethodE([Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] out string[] a)
        {
            throw new NotImplementedException();
        }
    }

    public static class COMMarshalFactory 
    {
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public static object Create() 
        {
            return new COMMarshal();
        }
    }
}
