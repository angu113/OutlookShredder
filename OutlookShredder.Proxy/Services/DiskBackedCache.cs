using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Two-level cache utility: L1 = volatile in-memory reference, L2 = JSON file on disk.
/// Disk writes are atomic (write to .tmp, then rename). All exceptions are swallowed
/// and logged so a corrupt/missing cache file never brings down the caller.
/// </summary>
public sealed class DiskBackedCache<T>
{
    private readonly string _path;
    private readonly ILogger _log;
    private readonly JsonSerializerOptions _opts;

    public DiskBackedCache(string cacheDir, string name, ILogger log, JsonSerializerOptions? opts = null)
    {
        _path = Path.Combine(cacheDir, $"{name}.json");
        _log  = log;
        _opts = opts ?? new JsonSerializerOptions();
        try { Directory.CreateDirectory(cacheDir); } catch { }
    }

    /// <summary>Reads and deserializes the disk cache. Returns default on any error.</summary>
    public T? TryLoad()
    {
        try
        {
            if (!File.Exists(_path)) return default;
            var json = File.ReadAllText(_path, Encoding.UTF8);
            return JsonSerializer.Deserialize<T>(json, _opts);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[DiskCache] Failed to load {Path}", _path);
            return default;
        }
    }

    /// <summary>Serializes <paramref name="data"/> and writes atomically via temp-then-rename.</summary>
    public async Task SaveAsync(T data)
    {
        var tmp = _path + ".tmp";
        try
        {
            var json = JsonSerializer.Serialize(data, _opts);
            await File.WriteAllTextAsync(tmp, json, Encoding.UTF8);
            File.Move(tmp, _path, overwrite: true);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[DiskCache] Failed to save {Path}", _path);
            try { File.Delete(tmp); } catch { }
        }
    }
}

/// <summary>
/// Deserializes JSON object? values to native .NET primitives (string, long, double, bool,
/// null, List&lt;object?&gt;, Dictionary&lt;string, object?&gt;) instead of JsonElement wrappers.
/// Required for correct round-trip of Dictionary&lt;string, object?&gt; through System.Text.Json.
/// On write, delegates to the runtime type's own serializer.
/// </summary>
internal sealed class ObjectConverter : JsonConverter<object?>
{
    public static readonly ObjectConverter Instance = new();

    public override object? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) =>
        reader.TokenType switch
        {
            JsonTokenType.True        => true,
            JsonTokenType.False       => false,
            JsonTokenType.Null        => null,
            JsonTokenType.Number      => reader.TryGetInt64(out var l) ? (object)l : reader.GetDouble(),
            JsonTokenType.String      => reader.GetString(),
            JsonTokenType.StartArray  => ReadArray(ref reader, options),
            JsonTokenType.StartObject => ReadObject(ref reader, options),
            _                         => throw new JsonException($"Unexpected token {reader.TokenType}")
        };

    private static List<object?> ReadArray(ref Utf8JsonReader reader, JsonSerializerOptions options)
    {
        var list = new List<object?>();
        while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
            list.Add(JsonSerializer.Deserialize<object?>(ref reader, options));
        return list;
    }

    private static Dictionary<string, object?> ReadObject(ref Utf8JsonReader reader, JsonSerializerOptions options)
    {
        var dict = new Dictionary<string, object?>(StringComparer.Ordinal);
        while (reader.Read() && reader.TokenType != JsonTokenType.EndObject)
        {
            var key = reader.GetString()!;
            reader.Read();
            dict[key] = JsonSerializer.Deserialize<object?>(ref reader, options);
        }
        return dict;
    }

    public override void Write(Utf8JsonWriter writer, object? value, JsonSerializerOptions options)
    {
        if (value is null) { writer.WriteNullValue(); return; }
        JsonSerializer.Serialize(writer, value, value.GetType(), options);
    }
}
