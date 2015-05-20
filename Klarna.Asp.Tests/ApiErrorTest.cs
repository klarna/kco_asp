﻿#region Copyright Header
// ----------------------------------------------------------------------------
// <copyright file="RecurringStatusTest.cs" company="Klarna AB">
//     Copyright 2015 Klarna AB
//
//     Licensed under the Apache License, Version 2.0 (the "License");
//     you may not use this file except in compliance with the License.
//     You may obtain a copy of the License at
//
//         http://www.apache.org/licenses/LICENSE-2.0
//
//     Unless required by applicable law or agreed to in writing, software
//     distributed under the License is distributed on an "AS IS" BASIS,
//     WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//     See the License for the specific language governing permissions and
//     limitations under the License.
// </copyright>
// <author>Klarna Support: support@klarna.com</author>
// <link>http://developers.klarna.com/</link>
// ----------------------------------------------------------------------------
#endregion
namespace Klarna.Asp.Tests
{
    using NUnit.Framework;

    /// <summary>
    /// Tests the ApiError class.
    /// </summary>
    [TestFixture]
    public class ApiErrorTest : TestCase
    {
        /// <summary>
        /// Runs the ASP unit tests for ApiError class.
        /// </summary>
        [Test]
        public void ApiErrorTests()
        {
            RunAspTests("ApiErrorTest");
        }
    }
}
