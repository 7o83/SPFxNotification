import * as React from 'react';
import styles from './SpFxNotificationList.module.scss';
import { ISpFxNotificationListProps } from './ISpFxNotificationListProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as IO from 'socket.io-client';
import { Spinner, MessageBar, AnimationStyles, IMessageBarStyles } from 'office-ui-fabric-react';

interface State {
  loading: boolean;
  notificationMsg: string;
  isMsgDisp: boolean;
  iMessageBarStyles: IMessageBarStyles;
}

//メッセージバーのアニメーション設定
const tmpIMessageBarStyles: IMessageBarStyles = {};

const openIMessageBarStyles: IMessageBarStyles = {};
openIMessageBarStyles.root = AnimationStyles.fadeIn200;

const closeIMessageBarStyles: IMessageBarStyles = {};
closeIMessageBarStyles.root = AnimationStyles.fadeOut200;


export default class SpFxNotificationList extends React.Component<ISpFxNotificationListProps, State> {

  //初期値の設定
  constructor(props: ISpFxNotificationListProps) {
    super(props);
    this.state = {
      loading: true, notificationMsg: 'No Message',
      isMsgDisp: false, iMessageBarStyles: tmpIMessageBarStyles
    };
  }

  //パーツ表示時の処理
  public async componentDidMount() {

    //socketioのエンドポイント取得
    const endpointUriSocketIO =
    `${location.protocol}//${location.hostname}/_api/v2.0` + 
    `/sites/${location.hostname}:${this.props.webPartContext.pageContext.web.serverRelativeUrl}:` + 
    `/lists/${this.props.noticeListName}/subscriptions/socketIo`;

    const res: SPHttpClientResponse =
      await this.props.webPartContext.spHttpClient.get(endpointUriSocketIO, SPHttpClient.configurations.v1);
    const resJson: any = await res.json();

    //socketioを使って変更通知の購読開始
    const socket = IO(resJson.notificationUrl, {
      transports: ['websocket']
    });

    //変更通知購読開始時の処理
    socket.on('connect', () => {
      console.log('Connect!', socket.id);
      this.setState({
        loading: false,
        notificationMsg: this.props.noticeListName + ' の更新を通知します ', isMsgDisp: true,
        iMessageBarStyles: openIMessageBarStyles
      });
    });

    socket.on('connect_error', (error: any) => console.log('Error!', error));
    socket.on('connect_timeout', (data: any) => console.log('Timeout!', data));

    //変更通知受信時の処理
    socket.on('notification', async (data: any) => {
      console.log('Notification!', data);

      if (this.state.isMsgDisp) {
        this.setState({ isMsgDisp: false, iMessageBarStyles: closeIMessageBarStyles });
      }
      const dataJson = JSON.parse(data);

      setTimeout(() => {
        const nowDatetime: Date = new Date();
        this.setState({
          notificationMsg: this.props.noticeListName + ' が更新されました ' + nowDatetime.toLocaleString('ja'),
          isMsgDisp: true, iMessageBarStyles: openIMessageBarStyles
        });
      }
        , 200);
    });
  }

  //描画処理
  public render(): React.ReactElement<ISpFxNotificationListProps> {
    return (
      <div>
        {this.state.loading ? (
          <Spinner label='loading...' />
        ) : (
            <MessageBar
              styles={this.state.iMessageBarStyles}
              dismissButtonAriaLabel='Close'
              onDismiss={(e) => { this.setState({ isMsgDisp: false, iMessageBarStyles: closeIMessageBarStyles }); }}
            >
              {this.state.notificationMsg}
            </MessageBar>
          )}
      </div>
    );
  }
}
