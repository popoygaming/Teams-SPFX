import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { INews } from '../INews';
import { MSGraphClient } from '@microsoft/sp-http';
import { IHelloWorldState } from '../IHelloWorldState';
import { IMail } from '../IMail';

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  
  constructor(props) {
    super(props);
    this.state = {
      newsArray: new Array<INews>(),
      mailsArray: new Array<IMail>(),
      profile: 
        { 
          displayName:"", 
          givenName:"", 
          mail:"" 
        }
    };
  }

  componentDidMount(){
    this.getProfile();
    this.getMyMails();
    this.getListItems();
  }

  public render(): React.ReactElement<IHelloWorldProps> {

    let currentLocation: string = (this.props.teamsContext) 
    ? `Teams: ${this.props.teamsContext.teamName}`
    : `Site Collection: ${this.props.webpartContext.pageContext.web.title}`;

    return (
      <div className={ styles.helloWorld }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to {escape(currentLocation)}!</span>
              
              <div>
                <div className={ styles.subTitle }>User Profile:</div>
                <ul>
                  <li>Name: {escape(this.state.profile.displayName)}</li>
                  <li>Mail: {escape(this.state.profile.mail)}</li>
                </ul>
              </div>
              
              <div>
                <div className={ styles.subTitle }>News List</div>
                <ul>
                  {( this.state.newsArray !== undefined && this.state.newsArray.length > 0) ?
                            this.state.newsArray.map((news, i) => {
                                return (
                                  <li>{news.Title}</li>
                                );
                              })
                    : <div><p><b>You have zero items in List</b></p></div>}
                </ul>
              </div>

              <div>
                <div className={ styles.subTitle }>Outlook Messages</div>
                <ul>
                  {( this.state.mailsArray !== undefined && this.state.mailsArray.length > 0) ?
                            this.state.mailsArray.map((mail, i) => {
                                return (
                                  <li>{mail.subject}</li>
                                );
                              })
                  : <div><p><b>You have zero items in Inbox</b></p></div>}
                  <a className={styles["custom-link"]} href="https://outlook.office.com/mail/inbox" target="_blank">View All Messages</a>
                </ul>
              </div>

            </div>
          </div>
        </div>
      </div>
    );
  }

  private getListItems() {
    let newsListId: string = "39ea750d-c5dd-495a-be19-2de4087af599";
    this.props.webpartContext.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void  =>{
      client
        .api('/sites/' + this.props.webpartContext.pageContext.site.id + '/lists/' + newsListId + '/items?expand=fields')
        .get ((error, response: any, rawResponse?: any)=>{
          
          let newsArray: Array<INews> = new Array<INews>();

          response.value.map((item: any) => {
            newsArray.push(item.fields);
            });
            
          this.setState({newsArray: newsArray});
        });
    });
  }

  private getProfile(){
    this.props.webpartContext.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/me')
          .get((error, response: any, rawResponse?: any) => {

            this.setState({profile: response});
        });
      });
  }

  private getMyMails() {
	    this.props.webpartContext.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void =>{
      client
        .api('/me/messages')
        .get ((error, mailResponse: any, rawResponse?: any)=>{
          
          var mailsArray: Array<IMail> = new Array<IMail>();
          mailResponse.value.map((item: IMail) => {
            mailsArray.push(item);
            });
            
          this.setState({mailsArray: mailsArray});
          // console.log(this.state.mailsArray);
        });
      });
  }
}
