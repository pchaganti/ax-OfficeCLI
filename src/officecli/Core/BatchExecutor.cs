// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Public entry point for running a batch against an already-open
/// <see cref="IDocumentHandler"/> without going through a CLI process or
/// resident pipe — for embedders that link officecli directly (in-process
/// hosts). Reuses the same batch-item dispatch and JSON
/// envelope shape as the CLI `batch` command and the Node/Python SDKs, so a
/// caller sees one consistent protocol regardless of transport.
/// </summary>
public static class BatchExecutor
{
    /// <param name="itemsJson">A JSON array of batch items — the same shape documented
    /// for the CLI `batch` command and the `BatchItem` SDK type.</param>
    /// <returns>A JSON array of per-item results (the same shape as the CLI's `--json` batch output).</returns>
    public static string ExecuteJson(IDocumentHandler handler, string itemsJson, bool stopOnError = false)
    {
        var items = JsonSerializer.Deserialize(itemsJson, BatchJsonContext.Default.ListBatchItem) ?? new List<BatchItem>();
        var results = CommandBuilder.RunNonResidentBatch(handler, items, stopOnError, json: true);
        return JsonSerializer.Serialize(results, BatchJsonContext.Default.ListBatchResult);
    }

    /// <summary>
    /// Mirrors the SDKs' <c>send(item)</c> (as distinct from <c>batch(items)</c>):
    /// runs ONE batch-shaped item and returns its own envelope directly — no
    /// index/success wrapper, and a failure throws instead of being captured
    /// per-item. Same underlying per-item dispatch as <see cref="ExecuteJson"/>,
    /// just not run through the batch driver's list wrapping/error-catching.
    /// </summary>
    /// <param name="itemJson">A single batch item object — the same shape as one
    /// entry in <see cref="ExecuteJson"/>'s array, or the SDKs' `send(item)` argument.</param>
    /// <returns>That item's own result envelope (e.g. `get`'s `{matches, results}` shape).</returns>
    public static string SendJson(IDocumentHandler handler, string itemJson)
    {
        var item = JsonSerializer.Deserialize(itemJson, BatchJsonContext.Default.BatchItem)
            ?? throw new ArgumentException("send: empty item");
        return CommandBuilder.ExecuteBatchItem(handler, item, json: true);
    }
}
