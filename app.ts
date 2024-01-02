import { SheetManager } from "./SheetManager";
import { TriggersManager } from "./TriggersManager";
import { GoogleSheetAPI } from "./GoogleSheetAPI";

function getSheetManager() {
    return SheetManager;
}

function getTriggersManager() {
    return TriggersManager;
}

function getGoogleSheetApi() {
    return GoogleSheetAPI;
}
