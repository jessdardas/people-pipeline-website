<%@ WebHandler Language="C#" Class="PeoplePipelineSettings" %>
// PEOPLE PIPELINE - shared settings for the website (the New DB date), saved in App_Data/settings.json.
//   GET  api/settings.ashx                                   -> the settings (everyone)
//   POST api/settings.ashx {"password":"…","newDbFrom":"2026-07-01"}   -> saves them (admin only)
//   POST api/settings.ashx {"password":"…","check":true}              -> only checks the password
// The password itself is not stored, only its fingerprint (the same one as in the Apps Script project).
// Needs: IIS with ASP.NET 4.x, and write permission for the app pool on the App_Data folder (see README).
using System;
using System.IO;
using System.Text;
using System.Web;
using System.Security.Cryptography;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

public class PeoplePipelineSettings : IHttpHandler
{
    const string AdminHash = "fcc5a0fbe0d3b33b95ce47e89c5ae92a24055ba5fb8833ad0e4c664703c0bbce";
    const string Defaults = "{\"newDbFrom\":\"2026-07-01\"}";

    public bool IsReusable { get { return false; } }

    public void ProcessRequest(HttpContext ctx)
    {
        string file = ctx.Server.MapPath("~/App_Data/settings.json");
        ctx.Response.ContentType = "application/json";
        ctx.Response.Cache.SetCacheability(HttpCacheability.NoCache);

        if (ctx.Request.HttpMethod != "POST")
        {
            ctx.Response.Write(File.Exists(file) ? File.ReadAllText(file) : Defaults);
            return;
        }

        var js = new JavaScriptSerializer();
        Dictionary<string, object> req;
        using (var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8))
        {
            try { req = js.Deserialize<Dictionary<string, object>>(reader.ReadToEnd()); }
            catch { req = null; }
        }
        if (req == null) { Fail(ctx, 400, "Bad request"); return; }

        string pw = req.ContainsKey("password") ? Convert.ToString(req["password"]) : "";
        if (Hash(pw) != AdminHash) { Fail(ctx, 403, "Wrong password"); return; }
        if (req.ContainsKey("check")) { ctx.Response.Write("{\"ok\":true}"); return; }

        string d = req.ContainsKey("newDbFrom") ? Convert.ToString(req["newDbFrom"]) : "";
        if (!Regex.IsMatch(d ?? "", @"^\d{4}-\d{2}-\d{2}$")) { Fail(ctx, 400, "Pick a valid date"); return; }

        long now = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds;
        string json = js.Serialize(new Dictionary<string, object> { { "newDbFrom", d }, { "savedAt", now } });
        Directory.CreateDirectory(Path.GetDirectoryName(file));
        File.WriteAllText(file, json, Encoding.UTF8);
        ctx.Response.Write(json);
    }

    static void Fail(HttpContext ctx, int status, string msg)
    {
        ctx.Response.StatusCode = status;
        ctx.Response.TrySkipIisCustomErrors = true;
        ctx.Response.Write("{\"error\":\"" + msg + "\"}");
    }

    static string Hash(string pw)
    {
        using (var sha = SHA256.Create())
        {
            var b = sha.ComputeHash(Encoding.UTF8.GetBytes("people-pipeline:" + pw));
            var sb = new StringBuilder();
            foreach (var x in b) sb.Append(x.ToString("x2"));
            return sb.ToString();
        }
    }
}
