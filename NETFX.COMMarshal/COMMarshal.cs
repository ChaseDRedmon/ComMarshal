using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
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
            [return: MarshalAs(UnmanagedType.BStr)] get;
            [param: MarshalAs(UnmanagedType.BStr)] set; 
        }
        
        int PropertyB 
        {
            [return: MarshalAs(UnmanagedType.I4)] get;
            [param: MarshalAs(UnmanagedType.I4)] set; 
        }
        
        bool PropertyC 
        {
            [return: MarshalAs(UnmanagedType.Bool)] get;
            [param: MarshalAs(UnmanagedType.Bool)] set;
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

        void InterfaceMarshal([In, MarshalAs(UnmanagedType.Interface)] Workbook book);
    }

    [ComVisible(true)]
    [Guid("a3ea8ead-cc9e-432b-a4e3-6810a2dee56e")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class COMMarshal : ICOMMarshal
    {
        /// <summary>
        /// A string class property
        /// </summary>
        public string PropertyA
        {
            [return: MarshalAs(UnmanagedType.BStr)] get;
            [param: MarshalAs(UnmanagedType.BStr)] set;
        }
        
        /// <summary>
        /// A integer class property
        /// </summary>
        public int PropertyB
        {
            [return: MarshalAs(UnmanagedType.I4)] get; 
            [param: MarshalAs(UnmanagedType.I4)] set;
        }
        
        /// <summary>
        /// A boolean class property
        /// </summary>
        public bool PropertyC
        {
            [return: MarshalAs(UnmanagedType.Bool)] get;
            [param: MarshalAs(UnmanagedType.Bool)] set;
        }
        
        /// <summary>
        /// A method that takes an integer
        /// </summary>
        /// <param name="C"></param>
        public void ArgumentA([In, MarshalAs(UnmanagedType.I4)] int A)
        {
            var a = Math.Pow(A, A);
            // Do something with a
        }
        
        /// <summary>
        /// A method that takes a single precision floating point value
        /// </summary>
        /// <param name="C"></param>
        public void ArgumentB([In, MarshalAs(UnmanagedType.R4)] float B)
        {
            var a = Math.Ceiling(B);
            // Do something with a
        }
        
        /// <summary>
        /// A method that takes a double precision floating point value
        /// </summary>
        /// <param name="C"></param>
        public void ArgumentC([In, MarshalAs(UnmanagedType.R8)] double C)
        {
            var a = Math.Floor(C);
            // Do something with a
        }
        
        /// <summary>
        /// A method that takes an array of strings
        /// </summary>
        /// <param name="E"></param>
        public void ArgumentD([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] D)
        {
            for (int i = 0; i < D.Length; i++)
            {
                // Do some processing with the string array
            }
        }
        
        /// <summary>
        /// A method that takes an array of integers
        /// </summary>
        /// <param name="E"></param>
        public void ArgumentE([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] ref int[] E)
        {
            for (int i = 0; i < E.Length; i++)
            {
                // Do some processing with the integer array
            }
        }
        
        /// <summary>
        /// A method that modifies the contents of the referenced integer array
        /// </summary>
        /// <param name="D"></param>
        public void MethodInOutA([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] ref string[] D)
        {
            for (int i = 0; i < D.Length; i++)
            {
                D[i] = D[i].ToUpper();
            }
        }
        
        /// <summary>
        /// A method that modifies the contents of the referenced integer array
        /// </summary>
        /// <param name="E"></param>
        public void MethodInOutB([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] ref int[] E)
        {
            for (int i = 0; i < E.Length; i++)
            {
                E[i] *= 2;
            }
        }
        
        /// <summary>
        /// A method that takes an integer and returns two integers - one from the return signature and one through the out parameter
        /// </summary>
        /// <param name="A"></param>
        /// <param name="B"></param>
        /// <returns></returns>
        [return: MarshalAs(UnmanagedType.I4)]
        public int MethodOutA([In, MarshalAs(UnmanagedType.I4)] int A, [Out, MarshalAs(UnmanagedType.I4)] out int B)
        {
            B = A * 2;
            return B * B;
        }
        
        /// <summary>
        /// Adds a worksheet to the workbook that called this method
        /// </summary>
        /// <code>
        /// Dim marshaller As COMMarshal
        /// Set marshaller = New COMMarshal
        /// marshaller.InterfaceMarshal ThisWorkbook
        /// </code>
        /// <param name="book"></param>
        public void InterfaceMarshal([In, MarshalAs(UnmanagedType.Interface)] Workbook book)
        {
            var sheet = book.Sheets.Add() as Worksheet;
            sheet.Name = "New Worksheet";
        }
        
        /// <summary>
        /// A simple .NET Method with no arguments and no return types
        /// </summary>
        public void MethodA()
        {
            Console.WriteLine("Hello from VBA");
        }
        
        /// <summary>
        /// A method that returns a boolean value
        /// </summary>
        /// <returns></returns>
        [return: MarshalAs(UnmanagedType.Bool)]
        public bool MethodB()
        {
            return true;
        }
        
        /// <summary>
        /// A method that returns an integer
        /// </summary>
        /// <returns></returns>
        [return: MarshalAs(UnmanagedType.I4)]
        public int MethodC()
        {
            return 10;
        }
        
        /// <summary>
        /// A method that returns an integer array via out statement
        /// </summary>
        /// <param name="a"></param>
        public void MethodD([Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)] out int[] a)
        {
            a = new[] { 1, 2, 3, 4, 5 };
        }
        
        /// <summary>
        /// A method that returns a string array via out statement
        /// </summary>
        /// <param name="a"></param>
        public void MethodE([Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_BSTR)] out string[] a)
        {
            a = new[] { "Hello", "from", "the", "other", "side" };
        }
    }
    
    /// <summary>
    /// ComMarshal Factory for creating COMMarshal objects as IDispatch types for consumption within VBA
    /// </summary>
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
