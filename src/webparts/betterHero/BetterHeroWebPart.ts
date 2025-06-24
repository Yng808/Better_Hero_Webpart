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
import Panzoom from '@panzoom/panzoom';

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
   sortOrder: number;
   zoom?: number;
   panX?: number;
   panY?: number;
}

// Acceptable image file types
const ImageExtensions = [
   ".apng", ".avif", ".bmp", ".gif", ".jpg", ".jpeg", ".jfif", ".ico",
   ".pjpeg", ".pjp", ".png", ".svg", ".svgz", ".tif", ".tiff", ".webp", ".xbm"
];

export default class BetterHeroWebPart extends BaseClientSideWebPart<IBetterHeroWebPartProps> {
   private imagesInfo: IImageInfo[];
   private form: Components.IForm;
   private panzoomInstances: Map<string, any> = new Map(); // Store panzoom instances

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

            // Add sortOrder to existing images that don't have it
            let needsUpdate = false;
            this.imagesInfo.forEach((imageInfo, index) => {
               if (imageInfo.sortOrder === undefined || imageInfo.sortOrder === null) {
                  imageInfo.sortOrder = index + 1; // Default sort order based on current position
                  needsUpdate = true;
               }
            });
            
            // Save updated images if sortOrder was added to null values
            if (needsUpdate) {
               this.properties.images = JSON.stringify(this.imagesInfo);
            }
            
            // Sort images by sortOrder property
            this.imagesInfo.sort((a, b) => a.sortOrder - b.sortOrder);
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

   // Render card. Pass input variables
   private renderCard(idx: number): HTMLElement {
      const imageInfo: IImageInfo = this.imagesInfo[idx];
      const isInEditMode = this.displayMode === DisplayMode.Edit;
      const elCard = document.createElement('div');
      //elCard.classList.add('col-md-6', 'col-lg-3');

      const imageId = `image-${this.instanceId}-${idx}`;

      const zoom = imageInfo.zoom || 1;
      const panX = imageInfo.panX || 0;
      const panY = imageInfo.panY || 0;
      const transformStyle = `transform: scale(${zoom}) translate(${panX / zoom}px, ${panY / zoom}px);`;

      elCard.innerHTML = `
            <a href="${isInEditMode ? '#' : imageInfo.url}">
             <div class="card bg-dark text-white ${styles.cardHideOverflow}">
                <div class="image-container" style="height: var(--hero-image-height); position: relative; overflow: hidden;">
                   <img id="${imageId}" 
                        src="${imageInfo.image}" 
                        class="card-img ${styles.leaderPhoto}" 
                        alt="${imageInfo.title}"
                        style="width: 100%; height: 100%; object-fit: cover; cursor: ${isInEditMode ? 'move' : 'pointer'}; transition: transform 0.3s ease; ${isInEditMode ? '' : transformStyle}">
                </div>
                <div class="card-img-overlay ${styles.cardImgOverlay}">
                   <h5 class="card-title ${styles.textTruncate}">${imageInfo.title}</h5>
                   <p class="card-text ${styles.textTruncate}">${imageInfo.subtitle || ''}</p>
                </div>    
             </div>             
          </a>

            ${isInEditMode ? `
             <div class="card card-buttons"></div>
             <div class="card-settings" style="padding: 10px; background: #f8f9fa; border-top: 1px solid #dee2e6;">
                <div class="row">
                   <div class="col-4">
                      <label style="font-size: 12px;">Zoom</label>
                      <input type="range" class="form-range zoom-slider" 
                             min="0.5" max="3" step="0.1" 
                             value="${imageInfo.zoom || 1}" 
                             data-image-id="${imageId}">
                      <span class="zoom-value" style="font-size: 11px;">${((imageInfo.zoom || 1) * 100).toFixed(0)}%</span>
                   </div>
                   <div class="col-4">
                      <label style="font-size: 12px;">Pan X</label>
                      <input type="range" class="form-range panx-slider" 
                             min="-100" max="100" step="1" 
                             value="${imageInfo.panX || 0}" 
                             data-image-id="${imageId}">
                      <span class="panx-value" style="font-size: 11px;">${imageInfo.panX || 0}</span>
                   </div>
                   <div class="col-4">
                      <label style="font-size: 12px;">Pan Y</label>
                      <input type="range" class="form-range pany-slider" 
                             min="-100" max="100" step="1" 
                             value="${imageInfo.panY || 0}" 
                             data-image-id="${imageId}">
                      <span class="pany-value" style="font-size: 11px;">${imageInfo.panY || 0}</span>
                   </div>
                </div>
                <div style="margin-top: 8px;">
                   <button class="btn btn-sm btn-outline-primary reset-btn" data-index="${idx}">Reset</button>
                   <button class="btn btn-sm btn-primary save-btn" data-index="${idx}">Save Position</button>
                </div>
             </div>
          ` : ''}
         `;

      if (isInEditMode) {
         // Setup Panzoom for edit mode
          setTimeout(() => {
             this.setupPanzoom(imageId, idx);
             this.setupSliderEvents(elCard, imageId, idx);
          }, 100);

         //render buttons
         Components.ButtonGroup({
            el: elCard.querySelector('.card-buttons') as HTMLElement,
            isSmall: true,
            buttons: [
               {
                  text: '✏️',
                  type: Components.ButtonTypes.OutlinePrimary,
                  onClick: () => {
                     // todo
                     this.renderForm(idx);
                  }
               },
               {
                  text: '🗑️',
                  type: Components.ButtonTypes.OutlineDanger,
                  onClick: () => {
                     // get element from array where image = selected one?, then delete
                     this.imagesInfo.splice(idx, 1);
                     this.properties.images = JSON.stringify(this.imagesInfo);
                     this.render();
                  }
               },
               {
                  text: '←',
                  type: Components.ButtonTypes.OutlineSecondary,
                  isDisabled: idx === 0, // Disable if it's the first item
                  onClick: () => {
                     this.moveItemUp(idx);
                  }
               },
               {
                  text: '→',
                  type: Components.ButtonTypes.OutlineSecondary,
                  isDisabled: idx === this.imagesInfo.length - 1, // Disable if it's the last item
                  onClick: () => {
                     this.moveItemDown(idx);
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
            {
               name: 'title',
               label: 'Title:',
               type: Components.FormControlTypes.TextField,
               required: true,
               errorMessage: 'Must specify a Title'
            },
            {
               name: 'sortOrder',
               label: 'Sort Order:',
               type: Components.FormControlTypes.TextField,
               required: false,
               errorMessage: 'Sort order must be a valid number'
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
                name: 'zoom',
                label: 'Default Zoom',
                type: Components.FormControlTypes.TextField,
                value: imageInfo?.zoom || 1
             },
             {
                name: 'panX',
                label: 'Default Pan X',
                type: Components.FormControlTypes.TextField,
                value: imageInfo?.panX || 0
             },
             {
                name: 'panY',
                label: 'Default Pan Y',
                type: Components.FormControlTypes.TextField,
                value: imageInfo?.panY || 0
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

   // Helper method to move item up in the list for the sort order
   private moveItemUp(idx: number): void {
      if (idx > 0) {
         // Swap with the previous item
         const temp = this.imagesInfo[idx];
         this.imagesInfo[idx] = this.imagesInfo[idx - 1];
         this.imagesInfo[idx - 1] = temp;
         
         // Update sortOrder values
         this.imagesInfo[idx].sortOrder = idx + 1;
         this.imagesInfo[idx - 1].sortOrder = idx;
         
         // Save and re-render
         this.properties.images = JSON.stringify(this.imagesInfo);
         this.render();
      }
   }

   // Helper method to move item down in the list for the sort order
   private moveItemDown(idx: number): void {
      if (idx < this.imagesInfo.length - 1) {
         // Swap with the next item
         const temp = this.imagesInfo[idx];
         this.imagesInfo[idx] = this.imagesInfo[idx + 1];
         this.imagesInfo[idx + 1] = temp;
         
         // Update sortOrder values
         this.imagesInfo[idx].sortOrder = idx + 1;
         this.imagesInfo[idx + 1].sortOrder = idx + 2;
         
         // Save and re-render
         this.properties.images = JSON.stringify(this.imagesInfo);
         this.render();
      }

   }


    private setupPanzoom(imageId: string, idx: number): void {
       const imageElement = document.getElementById(imageId) as HTMLElement;
       if (!imageElement) return;
 
       const imageInfo = this.imagesInfo[idx];
       
       // Initialize Panzoom
       const panzoom = Panzoom(imageElement, {
          maxScale: 3,
          minScale: 0.5,
          contain: 'outside',
          cursor: 'move',
          // Apply saved settings
          startX: imageInfo.panX || 0,
          startY: imageInfo.panY || 0,
          startScale: imageInfo.zoom || 1
       });
 
       // Store instance for cleanup
       this.panzoomInstances.set(imageId, panzoom);
 
       // Handle zoom with mouse wheel
       imageElement.parentElement?.addEventListener('wheel', panzoom.zoomWithWheel);
 
       // Update sliders when panzoom changes
       imageElement.addEventListener('panzoomchange', (event: any) => {
          const { x, y, scale } = event.detail;
          this.updateSliderValues(imageId, scale, x, y);
          
          // Auto-save position changes
          this.imagesInfo[idx] = {
             ...this.imagesInfo[idx],
             zoom: scale,
             panX: Math.round(x),
             panY: Math.round(y)
          };
       });
    }
 
    private setupSliderEvents(cardElement: HTMLElement, imageId: string, idx: number): void {
       const panzoom = this.panzoomInstances.get(imageId);
       if (!panzoom) return;
 
       // Zoom slider
       const zoomSlider = cardElement.querySelector('.zoom-slider') as HTMLInputElement;
       zoomSlider?.addEventListener('input', (e) => {
          const zoom = parseFloat((e.target as HTMLInputElement).value);
          panzoom.zoom(zoom, { animate: true });
          
          // Update display value immediately
          const zoomValue = cardElement.querySelector('.zoom-value');
          if (zoomValue) zoomValue.textContent = `${(zoom * 100).toFixed(0)}%`;
       });
 
       // Pan X slider
       const panXSlider = cardElement.querySelector('.panx-slider') as HTMLInputElement;
       panXSlider?.addEventListener('input', (e) => {
          const panX = parseInt((e.target as HTMLInputElement).value);
          const currentY = this.imagesInfo[idx].panY || 0;
          panzoom.pan(panX, currentY, { animate: true });
          
          // Update display value immediately
          const panXValue = cardElement.querySelector('.panx-value');
          if (panXValue) panXValue.textContent = panX.toString();
       });
 
       // Pan Y slider
       const panYSlider = cardElement.querySelector('.pany-slider') as HTMLInputElement;
       panYSlider?.addEventListener('input', (e) => {
          const panY = parseInt((e.target as HTMLInputElement).value);
          const currentX = this.imagesInfo[idx].panX || 0;
          panzoom.pan(currentX, panY, { animate: true });
          
          // Update display value immediately
          const panYValue = cardElement.querySelector('.pany-value');
          if (panYValue) panYValue.textContent = panY.toString();
       });
 
       // Reset button
       const resetBtn = cardElement.querySelector('.reset-btn') as HTMLButtonElement;
       resetBtn?.addEventListener('click', () => {
          this.resetImagePosition(idx, panzoom, cardElement);
       });
 
       // Save button
       const saveBtn = cardElement.querySelector('.save-btn') as HTMLButtonElement;
       saveBtn?.addEventListener('click', () => {
          this.saveImageSettings(idx);
       });
    }
 
    private updateSliderValues(imageId: string, scale: number, x: number, y: number): void {
       // Find sliders directly by their data-image-id attribute
       const zoomSlider = document.querySelector(`[data-image-id="${imageId}"].zoom-slider`) as HTMLInputElement;
       const panXSlider = document.querySelector(`[data-image-id="${imageId}"].panx-slider`) as HTMLInputElement;
       const panYSlider = document.querySelector(`[data-image-id="${imageId}"].pany-slider`) as HTMLInputElement;
 
       // Find value display elements (they're siblings of the sliders)
       const zoomValue = zoomSlider?.parentElement?.querySelector('.zoom-value');
       const panXValue = panXSlider?.parentElement?.querySelector('.panx-value');
       const panYValue = panYSlider?.parentElement?.querySelector('.pany-value');
 
       // Update slider positions
       if (zoomSlider) zoomSlider.value = scale.toString();
       if (panXSlider) panXSlider.value = Math.round(x).toString();
       if (panYSlider) panYSlider.value = Math.round(y).toString();
 
       // Update value displays
       if (zoomValue) zoomValue.textContent = `${(scale * 100).toFixed(0)}%`;
       if (panXValue) panXValue.textContent = Math.round(x).toString();
       if (panYValue) panYValue.textContent = Math.round(y).toString();
    }

 
    private resetImagePosition(idx: number, panzoom: any, cardElement: Element): void {
       // Reset to default values
       panzoom.reset({ animate: true });
       
       // Update image info
       this.imagesInfo[idx] = {
          ...this.imagesInfo[idx],
          zoom: 1,
          panX: 0,
          panY: 0
       };
 
       // Update UI
       this.updateSliderValues(`image-${this.instanceId}-${idx}`, 1, 0, 0);
    }
 
    private saveImageSettings(idx: number): void {
       // Save to web part properties
       this.properties.images = JSON.stringify(this.imagesInfo);
       
       // Show confirmation (you can customize this)
       const saveBtn = document.querySelector(`[data-index="${idx}"].save-btn`) as HTMLButtonElement;
       if (saveBtn) {
          const originalText = saveBtn.textContent;
          saveBtn.textContent = 'Saved!';
          saveBtn.style.backgroundColor = '#28a745';
          
          setTimeout(() => {
             saveBtn.textContent = originalText;
             saveBtn.style.backgroundColor = '';
          }, 2000);
       }
    }

 
    // Add cleanup method
    public onDispose(): void {
       // Clean up all panzoom instances
       this.panzoomInstances.forEach(panzoom => {
          panzoom.destroy();
       });
       this.panzoomInstances.clear();
    }

}