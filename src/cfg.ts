import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/*
 * SharePoint Assets
 */

export const Configuration = Helper.SPConfig({
    // create SP library with additional fields for: sort number, card title, card text, hyperlink, webpart ID (so that the same
    // doc lib can be used for multiple hero web parts across the site collection.)
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Libraries.Main,
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
            },
            CustomFields: [
                {
                    name: "SortOrder",
                    title: "Sort Order",
                    type: Helper.SPCfgFieldType.Number,
                },
                {
                    name: "CardTitle",
                    title: "Card Title",
                    type: Helper.SPCfgFieldType.Text,
                },
                {
                    name: "CardText",
                    title: "Card Text",
                    type: Helper.SPCfgFieldType.Text,
                },
                {
                    name: "WebpartId",
                    title: "Webpart ID",
                    type: Helper.SPCfgFieldType.Text,
                },
            ],
        },
    ],
});


