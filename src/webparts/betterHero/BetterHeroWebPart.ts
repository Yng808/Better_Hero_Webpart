import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'BetterHeroWebPartStrings';
import { Modal, LoadingDialog } from 'dattatable';
import { Components, Helper } from 'gd-sprest-bs';

export interface IBetterHeroWebPartProps {
  images: string; //store as JSON string
}


interface IImageInfo {
  image: string;
  title: string;
  url: string;
}

// Acceptable image file types
const ImageExtensions = [
  ".apng", ".avif", ".bmp", ".gif", ".jpg", ".jpeg", ".jfif", ".ico",
  ".pjpeg", ".pjp", ".png", ".svg", ".svgz", ".tif", ".tiff", ".webp", ".xbm"
];


export default class BetterHeroWebPart extends BaseClientSideWebPart<IBetterHeroWebPartProps> {
  private imagesInfo: IImageInfo[];
  private form: Components.IForm;

  public render(): void {
    // convert the images property to an array
    if (this.properties.images) {
      try {
        this.imagesInfo = JSON.parse(this.properties.images);
      } catch (error) {
        this.imagesInfo = [];
      }
    }
    else {
      this.imagesInfo = [];
    }

    // render the images
    this.renderImages();

  }

  private renderImages(): void {
    // see if images don't exist
    if (this.imagesInfo.length === 0) {
      this.domElement.innerHTML = 'There are no images. Edit the web part properties and add images.';
      return;
    }


    for (let i = 0; i < this.imagesInfo.length; i++) {
      //let imagesInfo = this.imagesInfo[i];
    }

    this.domElement.innerHTML = `    
      Todo
    `;
  }

  // Determines if the image extension is valid
  private isImageFile(fileName: string): boolean {
    let isValid = false;

    // Parse the valid file extensions
    for (let i = 0; i < ImageExtensions.length; i++) {
      // See if this is a valid file extension
      if (fileName.endsWith(ImageExtensions[i])) {
        // Set the flag and break from the loop
        isValid = true;
        break;
      }
    }

    // Return the flag
    return isValid;
  }


  private renderForm(): void {
    // render the form
    this.form = Components.Form({
      el: Modal.BodyElement,
      controls: [
        //title, image, url
        {
          label: 'Title:',
          type: Components.FormControlTypes.TextField,
          required: true,
          errorMessage: 'Must specify a Title'
        },
        {
          label: 'Link/Url:',
          type: Components.FormControlTypes.TextField,
          required: true,
          errorMessage: 'Must specify a Link/Url'
        },
        {
          label: 'Image:',
          type: Components.FormControlTypes.TextField,
          required: true,
          errorMessage: 'Must specify a Link/Url',
          onControlRendered: (ctrl) => {
            // Set a tooltip
            Components.Tooltip({
              content: "Click to upload an image file.",
              target: ctrl.textbox.elTextbox
            });

            // Make this textbox read-only
            ctrl.textbox.elTextbox.readOnly = true;

            // Set a click event
            ctrl.textbox.elTextbox.addEventListener("click", () => {
              // Display a file upload dialog
              Helper.ListForm.showFileDialog(["image/*"]).then(file => {
                // Clear the value
                ctrl.textbox.setValue("");

                // Get the file name
                let fileName = file.name.toLowerCase();

                // Validate the file type
                if (this.isImageFile(fileName)) {
                  // Show a loading dialog
                  LoadingDialog.setHeader("Reading the File");
                  LoadingDialog.setBody("This will close after the file is converted...");
                  LoadingDialog.show();

                  // Convert the file
                  let reader = new FileReader();
                  reader.onloadend = () => {
                    // Set the value
                    ctrl.textbox.setValue(reader.result as string);

                    // Close the dialog
                    LoadingDialog.hide();
                  }
                  reader.readAsDataURL(file.src);
                } else {
                  // Display an error message
                  ctrl.updateValidation(ctrl.el, {
                    isValid: false,
                    invalidMessage: "The file must be a valid image file. Valid types: png, svg, jpg, gif"
                  });
                }
              });
            });


          }
        },

      ]
    })

  }


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }


    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneButton('', {
                  text: strings.AddImageFieldLabel,
                  onClick: () => {

                    //set header
                    Modal.clear();
                    Modal.setHeader('Add image link');

                    // display a form
                    this.renderForm();

                    // render footer actions : save and cancel buttons
                    Components.TooltipGroup({
                      el: Modal.FooterElement,
                      tooltips: [
                        {
                          content: 'save the image link',
                          btnProps: {
                            text: 'Save',
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                              // ensure form is valid
                              if (this.form.isValid()) {
                                // todo this.form.getvalues. Go back to form, fill out name properties. Proeprty called name. This.form.getvalues() returns object  with keys. key matches the name of the form control
                                // the value of it is the value of the form. Create an object, add it to IImageInfo, push the object into the array. After this, save it. JSON stringify it, then save web part property.
                                // It should re-render the form after. After form re-renders, fill out renderImages() method. Just do this as pure html. need to consider editing an item, cleaning up code, separate out the new methods.
                              }
                            }

                          }
                        },
                        {
                          content: 'Close the dialog',
                          btnProps: {
                            text: 'Cancel',
                            type: Components.ButtonTypes.OutlineInfo,
                            onClick: () => {
                              Modal.hide()
                            }

                          }
                        }
                      ]
                    })
                    Modal.show()
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
