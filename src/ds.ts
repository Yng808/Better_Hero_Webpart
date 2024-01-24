import { List } from "dattatable";
import Strings from "./strings";
import { Components, Types, Web } from "gd-sprest-bs";

// list item interface
export interface IBetterHeroListItem extends Types.SP.ListItem {
    FileLeafRef: string; // The name of the file.
    FileRef: string; // The relative URL of the file.
    CardTitle: string;
    CardText: string;
}
// data source

// initialize the application

// return a promise

// initialize doc library