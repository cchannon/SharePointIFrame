import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { IFrameProps, LocationOption, iframe } from "./iframeControl";

export class documentsIframe implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        let locations: LocationOption[] = [];
        for (let index in context.parameters.DocumentLocations.records) {
            locations.push({
                dropdownOption: {
                    key: context.parameters.DocumentLocations.records[index].getRecordId(), 
                    text: context.parameters.DocumentLocations.records[index].getFormattedValue("name")
                }, 
                id: context.parameters.DocumentLocations.records[index].getRecordId(), 
                type: context.parameters.DocumentLocations.getTargetEntityType()
            });
        }
        
        const props: IFrameProps = { dropdownOptions: locations, context: context };

        return React.createElement(
            iframe, props
        );
    }

    public getOutputs(): IOutputs {
        return { };
    }

    public destroy(): void {
    }
}
