﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.IO;
using Microsoft.Bot.StreamingExtensions.Payloads;

namespace Microsoft.Bot.StreamingExtensions.UnitTests.Mocks
{
    public class MockPayloadAssembler : PayloadAssembler
    {
        public MockPayloadAssembler(Guid id)
            : base(id)
        {
        }

        public override Stream CreatePayloadStream()
        {
            return new MemoryStream();
        }

        public override void Close()
        {
            base.Close();
        }

        public override void OnReceive(Header header, Stream stream, int contentLength)
        {
            base.OnReceive(header, stream, contentLength);
        }
    }
}
