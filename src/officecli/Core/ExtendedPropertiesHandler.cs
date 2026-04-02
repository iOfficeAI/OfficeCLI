// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using AP = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeCli.Core;

/// <summary>
/// Shared Extended Properties (app.xml) Get/Set logic for all document types.
/// </summary>
public static class ExtendedPropertiesHandler
{
    /// <summary>
    /// Populate Format dictionary with extended properties.
    /// </summary>
    public static void PopulateExtendedProperties(ExtendedFilePropertiesPart? propsPart, DocumentNode node)
    {
        var props = propsPart?.Properties;
        if (props == null) return;

        if (props.Template?.Text != null)
            node.Format["extended.template"] = props.Template.Text;
        if (props.Manager?.Text != null)
            node.Format["extended.manager"] = props.Manager.Text;
        if (props.Company?.Text != null)
            node.Format["extended.company"] = props.Company.Text;
        if (props.Application?.Text != null)
            node.Format["extended.application"] = props.Application.Text;
        if (props.ApplicationVersion?.Text != null)
            node.Format["extended.applicationVersion"] = props.ApplicationVersion.Text;
        if (props.Pages?.Text != null)
            node.Format["extended.pages"] = int.TryParse(props.Pages.Text, out var p) ? (object)p : props.Pages.Text;
        if (props.Words?.Text != null)
            node.Format["extended.words"] = int.TryParse(props.Words.Text, out var w) ? (object)w : props.Words.Text;
        if (props.Characters?.Text != null)
            node.Format["extended.characters"] = int.TryParse(props.Characters.Text, out var c) ? (object)c : props.Characters.Text;
        if (props.Lines?.Text != null)
            node.Format["extended.lines"] = int.TryParse(props.Lines.Text, out var l) ? (object)l : props.Lines.Text;
        if (props.Paragraphs?.Text != null)
            node.Format["extended.paragraphs"] = int.TryParse(props.Paragraphs.Text, out var para) ? (object)para : props.Paragraphs.Text;
        if (props.TotalTime?.Text != null)
            node.Format["extended.totalTime"] = int.TryParse(props.TotalTime.Text, out var t) ? (object)t : props.TotalTime.Text;
    }

    /// <summary>
    /// Try to Set an extended.* property. Returns true if handled.
    /// </summary>
    public static bool TrySetExtendedProperty(ExtendedFilePropertiesPart? propsPart, string key, string value)
    {
        if (propsPart == null) return false;
        propsPart.Properties ??= new AP.Properties();
        var props = propsPart.Properties;

        switch (key)
        {
            case "extended.template":
                (props.Template ??= new AP.Template()).Text = value;
                break;
            case "extended.manager":
                (props.Manager ??= new AP.Manager()).Text = value;
                break;
            case "extended.company":
                (props.Company ??= new AP.Company()).Text = value;
                break;
            default:
                return false;
        }

        props.Save();
        return true;
    }

    /// <summary>
    /// Get the ExtendedFilePropertiesPart, creating if necessary for Set operations.
    /// </summary>
    public static ExtendedFilePropertiesPart? GetOrCreateExtendedPart(object doc)
    {
        return doc switch
        {
            WordprocessingDocument w => w.ExtendedFilePropertiesPart ?? w.AddExtendedFilePropertiesPart(),
            SpreadsheetDocument s => s.ExtendedFilePropertiesPart ?? s.AddExtendedFilePropertiesPart(),
            PresentationDocument p => p.ExtendedFilePropertiesPart ?? p.AddExtendedFilePropertiesPart(),
            _ => null
        };
    }

    public static ExtendedFilePropertiesPart? GetExtendedPart(object doc)
    {
        return doc switch
        {
            WordprocessingDocument w => w.ExtendedFilePropertiesPart,
            SpreadsheetDocument s => s.ExtendedFilePropertiesPart,
            PresentationDocument p => p.ExtendedFilePropertiesPart,
            _ => null
        };
    }
}
