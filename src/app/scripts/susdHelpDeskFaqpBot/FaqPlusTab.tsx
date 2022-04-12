import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the faqPlusTab React component
 */
export interface IFaqPlusTabState extends ITeamsBaseComponentState {

}

/**
 * Properties for the faqPlusTab React component
 */
export interface IFaqPlusTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the faqPlus content page
 */
export class FaqPlusTab extends TeamsBaseComponent<IFaqPlusTabProps, IFaqPlusTabState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content="Welcome to the SUSDHelpDeskFAQP Bot bot page" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>
                            <Text content="TODO: Add you content here" />
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright SUSD" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
