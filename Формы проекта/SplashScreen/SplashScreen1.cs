﻿using System;
using DevExpress.XtraSplashScreen;

namespace Simplex_2
{
    public partial class SplashScreen1 : SplashScreen
    {
        public SplashScreen1()
        {
            InitializeComponent();
        }
        #region Overrides

        public override void ProcessCommand(Enum cmd, object arg)
        {
            base.ProcessCommand(cmd, arg);            
        }

        #endregion

        public enum SplashScreenCommand
        {
        }
    }
}