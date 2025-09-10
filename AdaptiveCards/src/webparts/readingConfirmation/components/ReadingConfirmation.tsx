import * as React from 'react';
import styles from './ReadingConfirmation.module.scss';
import type { IReadingConfirmationProps } from './IReadingConfirmationProps';
import { TeamsMessageCreator } from '../../adaptiveCardViewer/components/TeamsMessageCreator';

export default class ReadingConfirmation extends React.Component<IReadingConfirmationProps> {
  public render(): React.ReactElement<IReadingConfirmationProps> {
    const {
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${styles.readingConfirmation} ${hasTeamsContext ? styles.teams : ''}`}>
        <TeamsMessageCreator 
          context={this.props.context}
          onMessageCreated={(messageId: number) => {
            console.log(`Message created with ID: ${messageId}`);
          }}
        />
      </section>
    );
  }
}
