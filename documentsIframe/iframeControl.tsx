import { Dropdown, IDropdownOption, IStackItemStyles, IStackStyles, IStackTokens, Stack } from "@fluentui/react";
import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";

export interface IFrameProps {
    dropdownOptions: LocationOption[],
    context: ComponentFramework.Context<IInputs>
}

export interface LocationOption {
    dropdownOption: IDropdownOption,
    id: string
    type: string
}

const stackTokens: IStackTokens = {
    childrenGap: 5
}

const stackStyles: IStackStyles = {
    root: {
        width: "100%",
    },
};
  
const stackItemStyles: IStackItemStyles = {
    root: {
        display: 'flex',
        width: "100%",
        height: "600px"
    },
};

export const iframe: React.FC<IFrameProps> = ((props: IFrameProps) => {
    const [options, setOptions] = React.useState<IDropdownOption[]>([]);
    const [selectedItem, setSelectedItem] = React.useState<IDropdownOption>();
    const [selectedUrl, setSelectedUrl] = React.useState<string>();
    let map : [{ key: string, url: string }] | undefined;

    const buildurl = (location: LocationOption, id: string, type: string, url: string) => {
        if(type.toLocaleLowerCase() === "sharepointdocumentlocation"){
            props.context.webAPI.retrieveRecord(type, id, "").then(
                (success) => {
                    if(success._parentsiteorlocation_value !== success.sitecollectionid){
                        buildurl(location, success._parentsiteorlocation_value, "sharepointdocumentlocation", success.relativeurl + "/" + url);
                    }
                    else{
                        buildurl(location, success._parentsiteorlocation_value, "sharepointsite", success.relativeurl + "/" + url);
                    }
                },
                (error) => {
                    console.log(error);
                }
            )
        } else { //sharepointsite
            props.context.webAPI.retrieveRecord(type, id, "").then(
                (success) => {
                    if(typeof(map) === "undefined"){
                        map =[{ key: location.id, url: success.absoluteurl + "/" + url }]; 
                    }else{
                        if(!map.includes({ key: location.id, url: success.absoluteurl + "/" + url }))
                            map.push({ key: location.id, url: success.absoluteurl + "/" + url });
                    }
                    let opts = options;
                    if(!opts.includes(props.dropdownOptions.find(x => x.id === location.id)!.dropdownOption))
                        opts.push(props.dropdownOptions.find(x => x.id === location.id)!.dropdownOption);
                    setOptions(opts);
                    setSelectedItem(options[0]);
                    setSelectedUrl(map?.find(x => x.key === options[0].key)?.url);
                },
                (error) => {
                    console.log(error);
                }
            )
        }
    };

    const onChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption<any> | undefined, index?: number | undefined) => {
        setSelectedItem(option);
        setSelectedUrl(map?.find(x => x.key === option?.key)?.url);
    };

    if(props.dropdownOptions.length > 0){
        props.dropdownOptions.forEach(opt => {
            buildurl(opt, opt.id, opt.type, "");
        });
    }
    
    return(
        <Stack tokens={stackTokens} styles={stackStyles}>
            <Stack.Item align="baseline" >
                <Dropdown
                    label="Document Location"
                    selectedKey={selectedItem ? selectedItem.key : (options.length > 0 ? options[0].key : "")}
                    onChange={onChange}
                    options={options}
                />
            </Stack.Item>
            <Stack.Item align="auto" styles={stackItemStyles}>
                <iframe src={selectedUrl} title="SharePoint Documents" width={"100%"}/> 
            </Stack.Item>
        </Stack>
    )
})