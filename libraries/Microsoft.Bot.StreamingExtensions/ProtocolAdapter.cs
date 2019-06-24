﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.StreamingExtensions.Payloads;
using Microsoft.Bot.StreamingExtensions.PayloadTransport;
using Microsoft.Bot.StreamingExtensions.Utilities;

namespace Microsoft.Bot.StreamingExtensions
{
    internal class ProtocolAdapter
    {
        private readonly RequestHandler _requestHandler;
        private readonly object _handlerContext;
        private readonly IPayloadSender _payloadSender;
        private readonly IPayloadReceiver _payloadReceiver;
        private readonly IRequestManager _requestManager;
        private readonly SendOperations _sendOperations;
        private readonly IStreamManager _streamManager;
        private readonly PayloadAssemblerManager _assemblerManager;

        public ProtocolAdapter(RequestHandler requestHandler, IRequestManager requestManager, IPayloadSender payloadSender, IPayloadReceiver payloadReceiver, object handlerContext = null)
        {
            _requestHandler = requestHandler;
            _handlerContext = handlerContext;
            _requestManager = requestManager;
            _payloadSender = payloadSender;
            _payloadReceiver = payloadReceiver;

            _sendOperations = new SendOperations(_payloadSender);
            _streamManager = new StreamManager(OnCancelStream);
            _assemblerManager = new PayloadAssemblerManager(_streamManager, OnReceiveRequest, OnReceiveResponse);

            _payloadReceiver.Subscribe(_assemblerManager.GetPayloadStream, _assemblerManager.OnReceive);
        }

        public async Task<ReceiveResponse> SendRequestAsync(Request request, CancellationToken cancellationToken)
        {
            var requestId = Guid.NewGuid();

            // wait for the response befores sending the request , so that we dont hit the case
            // where response is received before request is sent. This would lead to no tcs
            // being signaled
            var responseTask = _requestManager.GetResponseAsync(requestId, cancellationToken);

            var requestTask = _sendOperations.SendRequestAsync(requestId, request);

            if (cancellationToken != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            await Task.WhenAll(requestTask, responseTask).ConfigureAwait(false);

            return await responseTask;
        }

        private async Task OnReceiveRequest(Guid id, ReceiveRequest request)
        {
            // request is done, we can handle it
            if (_requestHandler != null)
            {
                var response = await _requestHandler.ProcessRequestAsync(request, _handlerContext).ConfigureAwait(false);

                if (response != null)
                {
                    await _sendOperations.SendResponseAsync(id, response).ConfigureAwait(false);
                }
            }
        }

        private async Task OnReceiveResponse(Guid id, ReceiveResponse response)
        {
            // we received the response to something, signal it
            await _requestManager.SignalResponse(id, response).ConfigureAwait(false);
        }

        private void OnCancelStream(PayloadAssembler contentStreamAssembler)
        {
            Background.Run(() => _sendOperations.SendCancelStreamAsync(contentStreamAssembler.Id));
        }
    }
}
