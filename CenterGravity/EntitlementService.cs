using System;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace BBI.JD
{
    public class EntitlementStatus
    {
        public bool IsValid { get; set; }
        public string Message { get; set; }
        public DateTime CheckedAtUtc { get; set; }
        public bool FromCache { get; set; }
    }

    internal class EntitlementResponse
    {
        [JsonPropertyName("UserId")]
        public string UserId { get; set; }

        [JsonPropertyName("AppId")]
        public string AppId { get; set; }

        [JsonPropertyName("IsValid")]
        public bool IsValid { get; set; }

        [JsonPropertyName("Message")]
        public string Message { get; set; }
    }

    /// <summary>
    /// Checks the Autodesk App Store Entitlement API (trial / purchase status) and
    /// caches the last known result locally so the tool keeps working offline.
    ///
    /// This is the "honor system" check Autodesk's own architecture provides for
    /// desktop add-ins - Autodesk does not enforce anything inside the plugin
    /// itself, so this call, and what to do with its result, is entirely on us.
    /// A client-side check like this can be bypassed by a determined user; it is
    /// not meant to be tamper-proof, only to honour the trial in normal use.
    /// </summary>
    public static class EntitlementService
    {
        public const string AppId = "2914827711437845009";

        private static readonly string CacheFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "JDS", "CenterGravity", "entitlement.json");

        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(8) };

        // --- TEMPORARY DEBUG SWITCH -------------------------------------------
        // Set the CENTERGRAVITY_FORCE_TRIAL_STATE environment variable to "expired"
        // or "valid" to force that state without waiting on the real trial or the
        // network - no rebuild needed, just relaunch Revit after setting it, e.g.
        //   $env:CENTERGRAVITY_FORCE_TRIAL_STATE = "expired"
        // Unset it (or set it to anything else) to go back to the real check:
        //   Remove-Item Env:\CENTERGRAVITY_FORCE_TRIAL_STATE
        // Remove this block before shipping a store build.
        private const string ForceStateEnvVar = "CENTERGRAVITY_FORCE_TRIAL_STATE";
        // -------------------------------------------------------------------

        public static async Task<EntitlementStatus> CheckAsync(string userId)
        {
            string forced = Environment.GetEnvironmentVariable(ForceStateEnvVar);

            if (string.Equals(forced, "expired", StringComparison.OrdinalIgnoreCase))
            {
                return new EntitlementStatus
                {
                    IsValid = false,
                    Message = "Forced by " + ForceStateEnvVar + "=expired (debug).",
                    CheckedAtUtc = DateTime.UtcNow
                };
            }

            if (string.Equals(forced, "valid", StringComparison.OrdinalIgnoreCase))
            {
                return new EntitlementStatus
                {
                    IsValid = true,
                    Message = "Forced by " + ForceStateEnvVar + "=valid (debug).",
                    CheckedAtUtc = DateTime.UtcNow
                };
            }

            if (string.IsNullOrWhiteSpace(userId))
            {
                // Not signed in to an Autodesk account: nothing to check against.
                // Fall back to the last known result, or fail open on a fresh
                // install (the App Store requires apps to be usable immediately).
                return ReadCache() ?? new EntitlementStatus
                {
                    IsValid = true,
                    Message = "Not signed in to an Autodesk account.",
                    CheckedAtUtc = DateTime.UtcNow
                };
            }

            try
            {
                string url = string.Format(
                    "https://apps.exchange.autodesk.com/webservices/checkentitlement/?userid={0}&appid={1}",
                    Uri.EscapeDataString(userId), Uri.EscapeDataString(AppId));

                string json = await Http.GetStringAsync(url).ConfigureAwait(false);

                EntitlementResponse response = JsonSerializer.Deserialize<EntitlementResponse>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                EntitlementStatus status = new()
                {
                    IsValid = response?.IsValid ?? true,
                    Message = response?.Message ?? string.Empty,
                    CheckedAtUtc = DateTime.UtcNow,
                    FromCache = false
                };

                WriteCache(status);
                return status;
            }
            catch (Exception)
            {
                // Offline / service unavailable: use the last known result rather
                // than punishing the user for a network hiccup. If we have never
                // checked successfully, fail open.
                EntitlementStatus cached = ReadCache();

                if (cached != null)
                {
                    cached.FromCache = true;
                    return cached;
                }

                return new EntitlementStatus
                {
                    IsValid = true,
                    Message = "Entitlement check unavailable.",
                    CheckedAtUtc = DateTime.UtcNow
                };
            }
        }

        private static EntitlementStatus ReadCache()
        {
            try
            {
                if (!File.Exists(CacheFile))
                {
                    return null;
                }

                string json = File.ReadAllText(CacheFile);
                return JsonSerializer.Deserialize<EntitlementStatus>(json);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void WriteCache(EntitlementStatus status)
        {
            try
            {
                string dir = Path.GetDirectoryName(CacheFile);

                if (!string.IsNullOrEmpty(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                File.WriteAllText(CacheFile, JsonSerializer.Serialize(status));
            }
            catch (Exception)
            {
                // best effort - a failed cache write just means the next launch
                // fails open again if offline, which is the safe direction.
            }
        }
    }
}
