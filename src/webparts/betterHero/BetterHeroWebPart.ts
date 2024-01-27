import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
   type IPropertyPaneConfiguration,
   PropertyPaneSlider,
   PropertyPaneButton,
   PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'BetterHeroWebPartStrings';
import { Modal, LoadingDialog } from 'dattatable';
import { Components, Helper } from 'gd-sprest-bs';
import 'bootstrap/dist/css/bootstrap.min.css';
import styles from './BetterHeroWebPart.module.scss';
import { BetterHeroWebPartHelpers } from './BetterHeroWebPartHelpers';

export interface IBetterHeroWebPartProps {
   images: string; //store as JSON string
   opacity: number;
   cardCols: number;
   cardHeight: number;
   cardColor: string;
}

interface IImageInfo {
   image: string;
   title: string;
   url: string;
   subtitle: string;
   hoverText: string;
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
      //const root = document.querySelector(':root') as HTMLElement;
      const uniqueClassName = `betterHeroWebPart-${this.instanceId}`;
      this.domElement.classList.add(uniqueClassName);

      if (!this.properties.cardCols) {
         this.properties.cardCols = 4;
      }
      // see if opacity exists and then set it
      if (this.properties.opacity >= 0 && this.properties.opacity <= 100) {
         //const root = document.querySelector(':root') as HTMLElement;
         //root.style.setProperty('--hero-image-opacity', (this.properties.opacity / 100).toString());
         this.domElement.style.setProperty('--hero-image-opacity', (this.properties.opacity / 100).toString());
      }

      // see if height exists and then set it
      if (this.properties.cardHeight >= 200 && this.properties.cardHeight <= 400) {
         //root.style.setProperty('--hero-image-height', (this.properties.cardHeight).toString() + 'px');
         this.domElement.style.setProperty('--hero-image-height', `${this.properties.cardHeight}px`);
      }

      // see if color exists and then set it
      if (this.properties.cardColor) {
         const [red, green, blue] = BetterHeroWebPartHelpers.getRGBFromColor(this.properties.cardColor);
         //root.style.setProperty('--hero-image-background-red', red.toString());
         //root.style.setProperty('--hero-image-background-green', green.toString());
         //root.style.setProperty('--hero-image-background-blue', blue.toString());
         this.domElement.style.setProperty('--hero-image-background-red', red.toString());
         this.domElement.style.setProperty('--hero-image-background-green', green.toString());
         this.domElement.style.setProperty('--hero-image-background-blue', blue.toString());
      }

      // convert the images property to an array
      if (this.properties.images) {
         try {
            // convert the JSON stringified string back to it's original form, the array
            this.imagesInfo = JSON.parse(this.properties.images);
         } catch (error) {
            this.imagesInfo = [];
         }
      }
      else {
         this.imagesInfo = [];
      }

      // render the images
      this.renderCards();

   }

   // render card. pass in input variables
   private renderCard(idx: number): HTMLElement {
      const imageInfo: IImageInfo = this.imagesInfo[idx];
      const isInEditMode = this.displayMode === DisplayMode.Edit;
      const elCard = document.createElement('div');
      //elCard.classList.add('col-md-6', 'col-lg-3');

      elCard.innerHTML = `
            <a href="${isInEditMode ? '#' : imageInfo.url}">
               <div class="card bg-dark text-white" style="overflow: hidden;">
                  <img src="${imageInfo.image}" class="card-img ${styles.leaderPhoto}" alt="${imageInfo.title}">
                  <div class="card-img-overlay ${styles.cardImgOverlay}">
                        <h5 class="card-title ${styles.textTruncate}">${imageInfo.title}</h5>
                        <p class="card-text ${styles.textTruncate}">${imageInfo.subtitle || ''}</p>
                  </div>    
               </div>             
            </a>

            ${isInEditMode ? '<div class="card card-buttons"></div>' : ''}
         `;

      if (isInEditMode) {
         //render buttons
         Components.ButtonGroup({
            el: elCard.querySelector('.card-buttons') as HTMLElement,
            isSmall: true,
            buttons: [
               {
                  text: 'Edit',
                  type: Components.ButtonTypes.OutlinePrimary,
                  onClick: () => {
                     // todo
                     this.renderForm(idx);
                  }
               },
               {
                  text: 'Delete',
                  type: Components.ButtonTypes.OutlineDanger,
                  onClick: () => {
                     // get element from array where image = selected one?, then delete
                     this.imagesInfo.splice(idx, 1);
                     this.properties.images = JSON.stringify(this.imagesInfo);
                     this.render();
                  }
               }
            ]
         })
      }

      if (imageInfo.hoverText) {
         Components.Tooltip({
            content: imageInfo.hoverText,
            target: elCard.querySelector('.card') as HTMLElement,
            type: Components.TooltipTypes.LightBorder
         })
      }

      return elCard;
   }

   private renderCards(): void {
      this.domElement.innerHTML = '';

      // check if images don't exist. If no images, display a message to guide the user to upload.
      if (this.imagesInfo.length === 0) {
         this.domElement.innerHTML = 'There are no images. Edit the web part properties and add images.';
         return;
      }

      // bootstrap container to make it responsive
      const elContainer = document.createElement('div');
      elContainer.classList.add('container-fluid', 'my-2');
      if (BetterHeroWebPartHelpers.isFullWidthSection(this.domElement)) {
         elContainer.style.marginLeft = "11px";
      }
      this.domElement.appendChild(elContainer);

      const elRows = document.createElement('div');

      // use custom column layouts when greater than 6 because bootstrap doesn't handle it nicely
      if (this.properties.cardCols <= 6) {
         // add row-cols-{num}
         elRows.classList.add('row', 'g-2', 'row-cols-md-2', 'row-cols-sm-1', `row-cols-lg-${this.properties.cardCols}`);
      }
      else {
         elRows.classList.add('row', 'g-2');
      }

      elContainer.appendChild(elRows);

      // debugger;
      // make a bootstrap card for each image in the array
      for (let i = 0; i < this.imagesInfo.length; i++) {
         const elCol = document.createElement('div');

         if (this.properties.cardCols <= 6) {
            elCol.classList.add('col-xs-1');
         }
         else {
            elCol.classList.add(`${styles.customCol}`, `${BetterHeroWebPartHelpers.getCustomColumnLayoutClass(this.properties.cardCols)}`);

         }

         elRows.appendChild(elCol);

         const elCard = this.renderCard(i);
         elCol.appendChild(elCard);
      }

      // this.domElement.innerHTML = html;
      // to clear out any javascript do this.domElement.innerHTML = this.domElement.innerHTML
   }

   private renderForm(idx: number = -1): void {
      const imageInfo = this.imagesInfo[idx];
      //set header
      Modal.clear();
      Modal.setHeader('Add image link');

      // render the form
      this.form = Components.Form({
         el: Modal.BodyElement,
         value: imageInfo,
         controls: [
            //title, image, url
            {
               name: 'title',
               label: 'Title:',
               type: Components.FormControlTypes.TextField,
               required: true,
               errorMessage: 'Must specify a Title'
            },
            {
               name: 'url',
               label: 'Link/Url:',
               type: Components.FormControlTypes.TextField,
               required: true,
               errorMessage: 'Must specify a Link/Url'
            },
            {
               name: 'subtitle',
               label: 'Subtitle',
               type: Components.FormControlTypes.TextField,
               required: true,
               errorMessage: 'Must specify a subtitle'
            },
            {
               name: 'hoverText',
               label: 'Hover Text',
               type: Components.FormControlTypes.TextArea
            },
            {
               name: 'image',
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
                        const fileName = file.name.toLowerCase();

                        // Validate the file type
                        if (BetterHeroWebPartHelpers.isImageFile(fileName, ImageExtensions)) {
                           // Show a loading dialog
                           LoadingDialog.setHeader("Reading the File");
                           LoadingDialog.setBody("This will close after the file is converted...");
                           LoadingDialog.show();

                           // Convert the file
                           const reader = new FileReader();
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

      // render footer actions : save and cancel buttons
      Components.TooltipGroup({
         el: Modal.FooterElement,
         tooltips: [
            {
               content: 'save the image link',
               btnProps: {
                  text: imageInfo ? 'Update' : 'Save',
                  type: Components.ButtonTypes.OutlinePrimary,
                  onClick: () => {
                     // ensure form is valid
                     if (this.form.isValid()) {
                        // get the form values
                        const formValues = this.form.getValues() as IImageInfo;

                        if (imageInfo) {
                           this.imagesInfo[idx] = formValues;
                        }
                        else {
                           // add object to the array
                           this.imagesInfo.push(formValues);
                        }


                        // Stringify because web part properties can only hold strings, numbers, and booleans
                        this.properties.images = JSON.stringify(this.imagesInfo);

                        this.render();

                        Modal.hide();
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


                              // display a form
                              this.renderForm();


                           }
                        }),
                        PropertyPaneSlider('opacity', {
                           label: strings.OpacityFieldLabel,
                           min: 0,
                           max: 100,
                           showValue: true
                        }),
                        PropertyPaneSlider('cardCols', {
                           label: strings.CardColFieldLabel,
                           min: 1,
                           max: 12,
                           showValue: true
                        }),
                        PropertyPaneSlider('cardHeight', {
                           label: strings.CardHeightFieldLabel,
                           min: 200,
                           max: 400,
                           showValue: true
                        }),
                        PropertyPaneDropdown('cardColor', {
                           label: strings.CardColorFieldLabel,
                           options: [
                              { key: 'black', text: 'black' },
                              { key: 'brown', text: 'brown' },
                              { key: 'deepRed', text: 'deep dark red' },
                              { key: 'green', text: 'green' },
                              { key: 'purple', text: 'purple' }
                           ]
                        })
                     ]
                  }
               ]
            }
         ]
      };
   }

}
