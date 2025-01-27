﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TcUnit.TcUnit_Runner {
    public class MessageFilter : IOleMessageFilter {
        // Class containing the IOleMessageFilter
        // thread error-handling functions.

        // Start the filter.
        public static void Register() {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(newFilter, out oldFilter);
        }

        // Done with the filter, close it.
        public static void Revoke() {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        //
        // IOleMessageFilter functions.
        // Handle incoming thread requests.
        int IOleMessageFilter.HandleInComingCall(int dwCallType,
          System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
          lpInterfaceInfo) {
            Console.WriteLine("SERVERCALL_ISHANDLED");
            //Return the flag SERVERCALL_ISHANDLED.
            return 0;
        }

        // Thread call was rejected, so try again.
        int IOleMessageFilter.RetryRejectedCall(System.IntPtr
          hTaskCallee, int dwTickCount, int dwRejectType) {
            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER.
            {

                // Retry the thread call immediately if return >=0 & 
                // <100.
                //Console.WriteLine("SERVERCALL_RETRYLATER");
                return 99;
            }
            // Too busy; cancel call.
            Console.WriteLine("Too busy");
            return -1;
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
          int dwTickCount, int dwPendingType) {
            //Return the flag PENDINGMSG_WAITDEFPROCESS.
            Console.WriteLine("PENDINGMSG_WAITDEFPROCESS");
            return 2;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int
          CoRegisterMessageFilter(IOleMessageFilter newFilter, out
          IOleMessageFilter oldFilter);
    }

    [System.Runtime.InteropServices.ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }
}
