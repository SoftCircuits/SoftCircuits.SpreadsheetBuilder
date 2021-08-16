// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Specifies whether <see cref="SpreadsheetBuilder"/> throws an exception
    /// when saving a document that does not pass validation.
    /// </summary>
    public enum SaveValidationExceptions
    {
        /// <summary>
        /// Never throw an exception when saving a document that does not pass
        /// validation.
        /// </summary>
        None,

        /// <summary>
        /// Always throw an exception when saving a document that does not pass
        /// validation.
        /// </summary>
        Always,

        /// <summary>
        /// Throw an exception when saving a document that does not pass validation
        /// only when running in debug mode.
        /// </summary>
        DebugOnly
    }
}
