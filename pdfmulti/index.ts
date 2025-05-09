import { IInputs, IOutputs } from "./generated/ManifestTypes";
import html2pdf from 'html2pdf.js';

export class pdfmulti implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _htmlContent: string;
    private _previewDiv: HTMLDivElement;
    private _convertButton: HTMLButtonElement;

    constructor() {
        // Empty constructor
    }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;

        // Create container div
        const containerDiv = document.createElement("div");
        containerDiv.className = "pdfmulti-container";
        this._container.appendChild(containerDiv);

        // Create preview div
        this._previewDiv = document.createElement("div");
        this._previewDiv.className = "pdfmulti-preview";
        containerDiv.appendChild(this._previewDiv);

        // Create convert button
        this._convertButton = document.createElement("button");
        this._convertButton.className = "pdfmulti-button";
        this._convertButton.textContent = "Download PDF";
        this._convertButton.addEventListener("click", this.convertToPDF.bind(this));
        containerDiv.appendChild(this._convertButton);

        // Initial render if content exists
        if (context.parameters.htmlInput) {
            this._htmlContent = context.parameters.htmlInput.raw || '';
            this._previewDiv.innerHTML = this._htmlContent;
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Update the HTML content when the input property changes
        if (context.parameters.htmlInput) {
            this._htmlContent = context.parameters.htmlInput.raw || '';
            this._previewDiv.innerHTML = this._htmlContent;
        }
    }

    private async convertToPDF(): Promise<void> {
        if (!this._htmlContent) {
            return;
        }

        // Create a temporary container for PDF conversion
        const tempContainer = document.createElement('div');
        tempContainer.style.cssText = ''; // Reset any default styles
        tempContainer.innerHTML = this._htmlContent;

        // Ensure all images are loaded before conversion
        const images = tempContainer.getElementsByTagName('img');
        const imagePromises = Array.from(images).map(img => {
            return new Promise((resolve, reject) => {
                if (img.complete) {
                    resolve(null);
                } else {
                    img.onload = () => resolve(null);
                    img.onerror = reject;
                }
            });
        });

        try {
            // Wait for all images to load
            await Promise.all(imagePromises);

            // Extract title from HTML content or use default
            const titleMatch = this._htmlContent.match(/<title>(.*?)<\/title>/i);
            const fileName = titleMatch ? titleMatch[1] : 'document';

            // Configure html2pdf options
            const opt = {
                margin: [10, 10, 20, 10] as [number, number, number, number],
                filename: `${fileName}.pdf`,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { 
                    scale: 2,
                    useCORS: true,
                    logging: true,
                    removeContainer: true,
                    windowWidth: undefined
                },
                jsPDF: { 
                    unit: 'mm', 
                    format: 'a4', 
                    orientation: 'portrait' as 'portrait' | 'landscape',
                    putOnlyUsedFonts: true,
                    floatPrecision: 16
                },
                pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
            };

            // Create PDF instance and handle page numbers
            const worker = html2pdf().set(opt);
            
            // Generate PDF with page numbers
            await worker
                .from(tempContainer)
                .save()
                .then(() => {
                    // Clean up after successful conversion
                    tempContainer.remove();
                });

        } catch (error) {
            console.error('PDF conversion failed:', error);
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        // Remove event listeners
        this._convertButton.removeEventListener("click", this.convertToPDF.bind(this));
    }
}
