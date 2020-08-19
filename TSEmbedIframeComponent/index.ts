import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class TSEmbedIframeComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

        // Reference to Bing Map IFrame HTMLElement
        private _embedIframe: HTMLElement;

		// Reference to the control container HTMLDivElement
		// This element contains all elements of our custom control example
        private _container: HTMLDivElement;

        // Flag if control view has been rendered
        private _controlViewRendered: Boolean;
        
		/**
		 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
		 * Data-set values are not initialized here, use updateView.
		 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
		 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
		 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
		 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
		 */
        public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
        {
			this._container = container;
			this._controlViewRendered = false;
			
        }

		/**
		 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
		 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
		 */
        public updateView(context: ComponentFramework.Context<IInputs>)
        {

            if (!this._controlViewRendered)
			{
                this._controlViewRendered = true;
                this.renderIFrameElement();
            }

            let src:string = context.parameters.passSrc.raw!;

            this._embedIframe.setAttribute("src", src);
		}

        /** 
         * Render IFrame HTML Element that hosts the Bing Map and appends the IFrame to the control container 
         */
        private renderIFrameElement(): void
        {
            this._embedIframe = this.createIFrameElement();
            this._container.appendChild(this._embedIframe);
        }

         /** 
         * Helper method to create an IFrame HTML Element
         */
        private createIFrameElement(): HTMLElement
        {
			let iFrameElement:HTMLElement = document.createElement("iframe");
			iFrameElement.setAttribute("class", "TS_EmbedIframe");
			iFrameElement.setAttribute("width", "100%");
			iFrameElement.setAttribute("height", "400px");
            return iFrameElement
        }

		/** 
		 * It is called by the framework prior to a control receiving new data. 
		 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
		 */
        public getOutputs(): IOutputs
        {
         	// no-op: method not leveraged by this example custom control
            return {};
        }

		/** 
 		 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
		 * i.e. cancelling any pending remote calls, removing listeners, etc.
		 */
        public destroy()
        {
            // no-op: method not leveraged by this example custom control
        }
}