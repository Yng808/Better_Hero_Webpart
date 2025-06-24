import styles from './BetterHeroWebPart.module.scss';

export class BetterHeroWebPartHelpers {

    public static getRGBFromColor(color: string): [number, number, number] {
        let red = 0;
        let green = 0;
        let blue = 0;

        switch (color) {
            case 'purple':
                red = 81;
                green = 14;
                blue = 153;
                break;
            case 'green':
                red = 0;
                green = 130;
                blue = 61;
                break;
            case 'deepRed':
                red = 102;
                green = 0;
                blue = 0;
                break;
            case 'brown':
                red = 164;
                green = 74;
                blue = 51;
                break;
        }

        return [red, green, blue];
    }

    public static isFullWidthSection(element: HTMLElement | undefined): boolean {
        //let element: HTMLElement | null = this.domElement;

        while (element) {
            if (element.classList && element.classList.contains('CanvasZone--fullWidth')) {
                // Class found, this is a full-width section
                return true;
            }
            element = element.parentElement || undefined;
        }

        // Full-width section class not found
        return false;
    }

    public static getCustomColumnLayoutClass(cardCols: number): string {
        let customColumnLayoutClass: string = '';
        switch (cardCols) {
            case 7:
                customColumnLayoutClass = styles.customCol7;
                break;
            case 8:
                customColumnLayoutClass = styles.customCol8;
                break;
            case 9:
                customColumnLayoutClass = styles.customCol9;
                break;
            case 10:
                customColumnLayoutClass = styles.customCol10;
                break;
            case 11:
                customColumnLayoutClass = styles.customCol11;
                break;
            case 12:
                customColumnLayoutClass = styles.customCol12;
                break;
        }
        return customColumnLayoutClass;
    }

    // Determines if the image extension is valid
    public static isImageFile(fileName: string, allowedExtensions: string[]): boolean {
        let isValid = false;

        // Parse the valid file extensions
        for (let i = 0; i < allowedExtensions.length; i++) {
            // See if this is a valid file extension
            if (fileName.endsWith(allowedExtensions[i])) {
                // Set the flag and break from the loop
                isValid = true;
                break;
            }
        }

        // Return the flag
        return isValid;
    }

    
}