import * as React from 'react';
import styles from './ScwAdmin.module.scss';
import { IItemDetailProps } from './IItemDetailProps';
import { IItemDetailState } from './IItemDetailState';
import { MessageBar, MessageBarType, IStackProps, Stack, ActionButton, IIconProps, DefaultButton, TextField, PrimaryButton, autobind, Spinner } from 'office-ui-fabric-react';
import { HttpClientResponse, HttpClient, IHttpClientOptions, HttpClientConfiguration, ODataVersion, IHttpClientConfiguration, AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {QueueClient, QueueServiceClient } from "@azure/storage-queue";

const backIcon: IIconProps = { iconName: 'NavigateBack' };

export default class ItemDetail extends React.Component<IItemDetailProps, IItemDetailState>{
    constructor(props: IItemDetailProps) {
        super(props);

        this.state = {
            showConfirmScr: false,
            status: '',
            comments: '',
            isLoading: false,
            showEndScr: false,
            resultStr: '',
            showDetailScr: true,
        };

    }

    public render(): React.ReactElement<IItemDetailProps> {

        return (
            <div className={styles.container}>
                {this.state.showDetailScr &&
                    <div className={styles.row}>
                        <span className={styles.title}>Request Detail</span>
                        <ActionButton
                            className={styles.newHeaderLinkStyle}
                            iconProps={backIcon}
                            allowDisabledFocus
                            onClick={this.props.returnToMainPage}>
                            Back to list
                </ActionButton>
                        <br />
                        <br />

                        <TextField label="Space Name" readOnly defaultValue={this.props.spaceName} />
                        <br />
                        <TextField label="Space Description (EN)" multiline rows={5} readOnly defaultValue={this.props.descriptionEn} />
                        <br />
                        <TextField label="Space Description (FR)" readOnly multiline rows={5} defaultValue={this.props.descriptionFR} />
                        <br />
                        <TextField label="Team Purpose and Content" readOnly multiline rows={3} defaultValue={this.props.reason} />
                        <br />
                        <TextField label="SharePoint Site Url" readOnly multiline rows={1} defaultValue={this.props.url} />
                        <br />
                        <TextField label="Requester Email" readOnly multiline rows={1} defaultValue={this.props.requesterEmail} />
                        <br />

                        <PeoplePicker
                            context={this.props.context}
                            titleText="Owners"
                            personSelectionLimit={3}
                            groupName={""}
                            showHiddenInUI={false}
                            defaultSelectedUsers={this.props.owners}
                            disabled={true}
                            ensureUser={false} />
                        <br />
                        {this.props.itemStatus == 'Submitted' &&
                            <Stack horizontal verticalAlign="center" horizontalAlign="center">
                                <PrimaryButton text="Approve" style={{ marginRight: "20px" }} onClick={this._btnApproveClicked} allowDisabledFocus />
                                <PrimaryButton text="Reject" onClick={this._btnRejectClicked} allowDisabledFocus />
                            </Stack>
                        }
                    </div>
                }
                {this.state.showConfirmScr &&
                    <div className={styles.row}>
                        <span className={styles.title}>Confirm</span>
                        <br />
                        <TextField label="Comment (optional)" multiline rows={5} placeholder="Type a comment to send to the requestor" onChanged={this._onChangedComments} />
                        <br />

                        {this.state.isLoading &&
                            <Spinner label="Update item..." />

                        }
                        <br />
                        <Stack horizontal verticalAlign="center" horizontalAlign="center">
                            <PrimaryButton text="Back" style={{ marginRight: "20px" }} onClick={this._back} allowDisabledFocus />
                            <PrimaryButton text="Comfirm" onClick={this._callAzureFunction} allowDisabledFocus />
                        </Stack>

                    </div>
                }
                {this.state.showEndScr &&
                    <div className={styles.row}>
                        <span className={styles.title}>Item update {this.state.resultStr}.</span>
                        <br /> <br /> <br />
                        <Stack horizontal verticalAlign="center" horizontalAlign="center">
                            <PrimaryButton text="Close" onClick={this.props.returnToMainPage} />
                        </Stack>
                    </div>

                }

            </div>
        );
    }


    @autobind
    private _onChangedComments(comments: string): void {
        this.setState({
            comments: comments,
        });
    }

    @autobind
    private _btnRejectClicked(): void {
        this.setState({
            showConfirmScr: true,
            status: "Rejected",
            showDetailScr: false,
        });
    }

    @autobind
    private _btnApproveClicked(): void {
        this.setState({
            showConfirmScr: true,
            showDetailScr: false,
            status: "Approved",
        });
        console.log("key", this.props.itemId);
    }

    @autobind
    private _back(): void {
        this.setState({
            showConfirmScr: false,
            showDetailScr: true,

        });
    }


    protected functionUrl: string = "https://gxdccps-updatelist-fnc.azurewebsites.net/api/UpdateList?";
    protected funcUrl: string = "https://gxdccps-sitescreations-fnc.azurewebsites.net/api/CreateSiteQueue";
    protected emailQueueUrl: string ="https://gxdccps-emailnotification-fnc.azurewebsites.net/api/EmailQueue?";
    @autobind
    private _callAzureFunction(): void {

        console.log("key ", this.props.itemId);
        console.log("comments ", this.state.comments);
        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");
        const postOptions: IHttpClientOptions = {
            headers: requestHeaders,
            body: `
              {
                "name": 
                {
                  "key": "${this.props.itemId}",
                  "comments": "${this.state.comments}",      
                  "status": "${this.state.status}", 
                }
              }`
        };

        this.setState({
            isLoading: true,
        });
        // update status
        this.props.context.aadHttpClientFactory.getClient("b93d2cc6-1944-4e2a-a2db-b50e1a474268").then((client: AadHttpClient) => {
            client.post(this.functionUrl, AadHttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
                console.log(`Status code:`, response.status);
                console.log('respond is ', response.ok);
                if (response.ok) {
                    this.setState({
                        isLoading: false,
                        showEndScr: true,
                        showConfirmScr: false,
                        showDetailScr: false,
                        resultStr: 'completed'
                    });
                } else {
                    this.setState({
                        isLoading: false,
                        showEndScr: true,
                        showConfirmScr: false,
                        showDetailScr: false,
                        resultStr: 'failed'
                    });
                }

            });
        });

        // http trigger createsite function app
        if (this.state.status === "Approved"){
            // let mailNickname = this.props.url.substring(39);
            // let emailsStr = this.props.owners.toString();
            // console.log(emailsStr);

            const postQueue: IHttpClientOptions = {
                headers: requestHeaders,
                body: `
                {
                    "name": "${this.props.spaceName}",
                    "description": "${this.props.descriptionEn}-${this.props.descriptionFR}" ,
                    "mailNickname": "${this.props.url.substring(39)}",
                    "itemId": "${this.props.itemId}",
                    "emails": "${this.props.owners.toString()}",
                    "requesterName": "${this.props.requesterName}",
                    "requesterEmail": "${this.props.requesterEmail}"
                }`
            };
            this.props.context.aadHttpClientFactory.getClient("b93d2cc6-1944-4e2a-a2db-b50e1a474268").then((client: AadHttpClient) => {
                client.post(this.funcUrl, AadHttpClient.configurations.v1, postQueue).then((response: HttpClientResponse) => {
                    console.log(`Status code:`, response.status);
                    console.log('respond is ', response.ok);
                    console.log('send message successful.');

                });
            });
        }else{
            // http trigger sendStatusToQueue function app
            const postQueue: IHttpClientOptions = {
                headers: requestHeaders,
                body: `
                {
                    "name": "${this.props.spaceName}",
                    "status": "${this.state.status}",
                    "comments": "${this.state.comments}",
                    "requesterName": "${this.props.requesterName}",
                    "requesterEmail": "${this.props.requesterEmail}"
                }`
            };
            this.props.context.aadHttpClientFactory.getClient("b93d2cc6-1944-4e2a-a2db-b50e1a474268").then((client: AadHttpClient) => {
                client.post(this.emailQueueUrl, AadHttpClient.configurations.v1, postQueue).then((response: HttpClientResponse) => {
                    console.log(`Status code:`, response.status);
                    console.log('respond is ', response.ok);
                    console.log('send reject message to queue successful.');
                    console.log(`requester Email`, this.props.requesterEmail);
                });
            });   
        }
    }

}