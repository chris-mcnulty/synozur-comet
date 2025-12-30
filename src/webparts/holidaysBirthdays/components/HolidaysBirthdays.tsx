import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import {
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  IconButton,
  Stack,
  Text,
  Image,
  ImageFit,
  PrimaryButton,
  getTheme
} from 'office-ui-fabric-react';
import { IEventOccurrence } from '../services/RecurrenceCalculator';
import { DataService } from '../services/DataService';
import { ListProvisioningService } from '../services/ListProvisioningService';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './HolidaysBirthdays.module.scss';
import { DisplayMode } from '../IHolidaysBirthdaysWebPartProps';

// Import default images
const defaultHolidayImage = require('../assets/default-holiday.jpg');
const defaultBirthdayImage = require('../assets/default-birthday.jpg');
const defaultChristmasImage = require('../assets/default-christmas.jpg');
const defaultIndependenceImage = require('../assets/default-independence.jpg');
const synozurLogo = require('../assets/synozur-logo.png');

export interface IHolidaysBirthdaysProps {
  context: WebPartContext;
  daysDefault: number;
  daysExpanded: number;
  listTitle: string;
  showImages: boolean;
  showTypeBadges: boolean;
  allowListProvisioning: boolean;
  displayMode: DisplayMode;
  showFooter: boolean;
}

const HolidaysBirthdays: React.FunctionComponent<IHolidaysBirthdaysProps> = (props) => {
  const [occurrences, setOccurrences] = useState<IEventOccurrence[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [isExpanded, setIsExpanded] = useState<boolean>(false);

  // SharePoint Brand Center fonts are automatically available via CSS variables
  // The SCSS file uses --fontFamilyCustomFont100 for body text and --fontFamilyCustomFont1000 for headlines
  // These variables are set by SharePoint when a custom font package is applied to the site

  const initializePnP = useCallback(() => {
    sp.setup({
      spfxContext: props.context as any
    });
  }, [props.context]);

  const filterEventsByDisplayMode = useCallback((events: IEventOccurrence[]): IEventOccurrence[] => {
    if (props.displayMode === 'Both') {
      return events;
    }
    if (props.displayMode === 'Holidays') {
      return events.filter(e => e.eventType === 'Holiday');
    }
    if (props.displayMode === 'Birthdays') {
      return events.filter(e => e.eventType === 'Birthday');
    }
    return events;
  }, [props.displayMode]);

  const loadData = useCallback(async (maxDays: number) => {
    try {
      setLoading(true);
      setError(null);

      // Initialize PnP before any operations
      initializePnP();
      
      // Check if list exists and provision if needed
      if (props.allowListProvisioning) {
        const provisioningService = new ListProvisioningService(props.listTitle);
        const provisioningResult = await provisioningService.ensureListExists();
        
        if (!provisioningResult.success) {
          const errorMsg = provisioningResult.error || 'Unknown error';
          setError(`Unable to access or create the list "${props.listTitle}". Please ensure you have permissions to create lists and fields, or contact your administrator. Error: ${errorMsg}`);
          setLoading(false);
          return;
        }
      }

      // Load events
      const dataService = new DataService(props.listTitle);
      const events = await dataService.getEvents(maxDays);
      
      // Filter events based on display mode
      const filteredEvents = filterEventsByDisplayMode(events);
      
      setOccurrences(filteredEvents);
      setLoading(false);
    } catch (err) {
      console.error('Error loading data:', err);
      const errorMessage = err instanceof Error ? err.message : 'Unknown error occurred';
      
      // Check if this is a list not found error
      if (errorMessage.includes('does not exist') || errorMessage.includes('404') || errorMessage.includes('ItemNotFoundException')) {
        if (props.allowListProvisioning) {
          // If provisioning is enabled but we still get this error, provisioning likely failed
          setError(`The list "${props.listTitle}" does not exist and could not be created automatically. Please ensure you have "Manage Lists" permission, or contact your administrator.`);
        } else {
          setError(`The list "${props.listTitle}" does not exist. Please create the list or enable "Allow List Provisioning" in web part settings.`);
        }
      } else {
        setError(`Error loading events: ${errorMessage}`);
      }
      setLoading(false);
    }
  }, [props.context, props.listTitle, props.allowListProvisioning, initializePnP, filterEventsByDisplayMode]);

  useEffect(() => {
    initializePnP();
  }, [initializePnP]);

  useEffect(() => {
    const maxDays = isExpanded ? props.daysExpanded : props.daysDefault;
    loadData(maxDays);
  }, [isExpanded, props.daysDefault, props.daysExpanded, props.displayMode, loadData]);

  const getTitle = (): string => {
    switch (props.displayMode) {
      case 'Holidays':
        return 'Upcoming Holidays';
      case 'Birthdays':
        return 'Upcoming Birthdays';
      default:
        return 'Upcoming Events';
    }
  };

  const formatDate = (date: Date): string => {
    return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
  };

  const groupByMonth = (events: IEventOccurrence[]): { [key: string]: IEventOccurrence[] } => {
    const groups: { [key: string]: IEventOccurrence[] } = {};
    
    events.forEach(event => {
      const monthKey = event.date.toLocaleDateString('en-US', { year: 'numeric', month: 'long' });
      if (!groups[monthKey]) {
        groups[monthKey] = [];
      }
      groups[monthKey].push(event);
    });
    
    return groups;
  };

  const getTypeBadgeColor = (eventType: string): string => {
    const theme = getTheme();
    return eventType === 'Birthday' ? theme.palette.themePrimary : theme.palette.themeSecondary;
  };

  const getDefaultImage = (event: IEventOccurrence): string => {
    // Check for specific holidays first
    if (event.eventType === 'Holiday') {
      const month = event.date.getMonth(); // 0-indexed (0 = January, 11 = December)
      const day = event.date.getDate();
      
      // Check for Christmas (December 25)
      if (month === 11 && day === 25) {
        return defaultChristmasImage;
      }
      
      // Check for Independence Day (July 4)
      if (month === 6 && day === 4) {
        return defaultIndependenceImage;
      }
      
      // Fallback to generic holiday image
      return defaultHolidayImage;
    } else if (event.eventType === 'Birthday') {
      return defaultBirthdayImage;
    }
    
    // Fallback to holiday image if event type is unknown
    return defaultHolidayImage;
  };

  const getImageUrl = (event: IEventOccurrence): string | undefined => {
    // Use custom image if provided, otherwise use default based on event type and date
    return event.imageUrl || getDefaultImage(event);
  };

  const getListUrl = (): string => {
    return `${props.context.pageContext.web.absoluteUrl}/Lists/${props.listTitle}`;
  };

  if (loading && occurrences.length === 0) {
    return (
      <div className={styles.holidaysBirthdays}>
        <Spinner size={SpinnerSize.large} label="Loading events..." />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.holidaysBirthdays}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {error}
        </MessageBar>
        {!props.allowListProvisioning && (
          <Text className={styles.errorHint}>
            Tip: Enable "Allow List Provisioning" in web part settings to automatically create the list.
          </Text>
        )}
      </div>
    );
  }

  const groupedEvents = groupByMonth(occurrences);
  const monthKeys = Object.keys(groupedEvents).sort((a, b) => {
    const dateA = new Date(a);
    const dateB = new Date(b);
    return dateA.getTime() - dateB.getTime();
  });

  return (
    <div className={styles.holidaysBirthdays}>
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="xLarge" className={styles.title}>
            {getTitle()} {!isExpanded && `(Next ${props.daysDefault} days)`}
          </Text>
          <IconButton
            iconProps={{ iconName: isExpanded ? 'ChevronUp' : 'ChevronDown' }}
            title={isExpanded ? 'Show fewer events' : 'Show more events'}
            onClick={() => setIsExpanded(!isExpanded)}
            className={styles.toggleButton}
          />
        </Stack>

        {occurrences.length === 0 ? (
          <Stack tokens={{ childrenGap: 12 }} className={styles.emptyState}>
            <Text variant="medium">No upcoming events found.</Text>
            <PrimaryButton
              text="Add Events"
              href={getListUrl()}
              target="_blank"
            />
          </Stack>
        ) : (
          <div className={isExpanded ? styles.eventsContainerExpanded : styles.eventsContainer}>
            <Stack tokens={{ childrenGap: 24 }}>
              {monthKeys.map(monthKey => (
                <Stack key={monthKey} tokens={{ childrenGap: 12 }}>
                  <Text variant="large" className={styles.monthHeader}>
                    {monthKey}
                  </Text>
                  <Stack tokens={{ childrenGap: 8 }}>
                    {groupedEvents[monthKey].map((event, index) => (
                      <div key={`${event.date.getTime()}-${index}`} className={styles.eventItem}>
                        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                          {props.showImages && (
                            <Image
                              src={getImageUrl(event)}
                              alt={event.title}
                              width={48}
                              height={48}
                              imageFit={ImageFit.cover}
                              className={styles.eventImage}
                            />
                          )}
                          <Stack tokens={{ childrenGap: 4 }} grow>
                            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                              <Text variant="medium" className={styles.eventDate}>
                                {formatDate(event.date)}
                              </Text>
                              <Text variant="medium" className={styles.eventTitle}>
                                {event.title}
                              </Text>
                              {props.showTypeBadges && (
                                <span
                                  className={styles.typeBadge}
                                  style={{ backgroundColor: getTypeBadgeColor(event.eventType) }}
                                >
                                  {event.eventType}
                                </span>
                              )}
                            </Stack>
                            {event.notes && (
                              <Text variant="small" className={styles.eventNotes}>
                                {event.notes}
                              </Text>
                            )}
                          </Stack>
                        </Stack>
                      </div>
                    ))}
                  </Stack>
                </Stack>
              ))}
            </Stack>
          </div>
        )}
        {props.showFooter && (
          <div className={styles.footer}>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" horizontalAlign="center">
              <Image
                src={synozurLogo}
                alt="Synozur"
                width={128}
                height={128}
                imageFit={ImageFit.contain}
                className={styles.footerLogo}
              />
              <Text variant="small" className={styles.footerText}>
                Holidays & Birthdays • © {new Date().getFullYear()} The Synozur Alliance LLC. All Rights Reserved.
              </Text>
            </Stack>
          </div>
        )}
      </Stack>
    </div>
  );
};

export default HolidaysBirthdays;