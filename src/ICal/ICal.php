<?php
/**
 * This PHP class will read an ICS (`.ics`, `.ical`, `.ifb`) file, parse it and return an
 * array of its contents.
 *
 * PHP 5 (≥ 5.3.0)
 *
 * @author  Jonathan Goode <https://github.com/u01jmg3>
 * @license https://opensource.org/licenses/mit-license.php MIT License
 * @version 2.1.9
 */

namespace ICal;

use Carbon\Carbon;

class ICal
{
    // phpcs:disable Generic.Arrays.DisallowLongArraySyntax.Found

    const DATE_TIME_FORMAT        = 'Ymd\THis';
    const DATE_TIME_FORMAT_PRETTY = 'F Y H:i:s';
    const ICAL_DATE_TIME_TEMPLATE = 'TZID=%s:';
    const RECURRENCE_EVENT        = 'Generated recurrence event';
    const SECONDS_IN_A_WEEK       = 604800;
    const TIME_FORMAT             = 'His';
    const TIME_ZONE_UTC           = 'UTC';
    const UNIX_FORMAT             = 'U';
    const UNIX_MIN_YEAR           = 1970;

    /**
     * Tracks the number of alarms in the current iCal feed
     *
     * @var integer
     */
    public $alarmCount = 0;

    /**
     * Tracks the number of events in the current iCal feed
     *
     * @var integer
     */
    public $eventCount = 0;

    /**
     * Tracks the free/busy count in the current iCal feed
     *
     * @var integer
     */
    public $freeBusyCount = 0;

    /**
     * Tracks the number of todos in the current iCal feed
     *
     * @var integer
     */
    public $todoCount = 0;

    /**
     * The value in years to use for indefinite, recurring events
     *
     * @var integer
     */
    public $defaultSpan = 2;

    /**
     * Enables customisation of the default time zone
     *
     * @var string
     */
    public $defaultTimeZone;

    /**
     * The two letter representation of the first day of the week
     *
     * @var string
     */
    public $defaultWeekStart = 'MO';

    /**
     * Toggles whether to skip the parsing of recurrence rules
     *
     * @var boolean
     */
    public $skipRecurrence = false;

    /**
     * Toggles whether to use time zone info when parsing recurrence rules
     *
     * @var boolean
     */
    public $useTimeZoneWithRRules = false;

    /**
     * Toggles whether to disable all character replacement.
     *
     * @var boolean
     */
    public $disableCharacterReplacement = false;

    /**
     * Toggles whether to replace (non-CLDR) Windows time zone IDs with their IANA equivalent.
     *
     * @var boolean
     */
    public $replaceWindowsTimeZoneIds = false;

    /**
     * With this being non-null the parser will ignore all events more than roughly this many days after now.
     *
     * @var integer
     */
    public $filterDaysBefore = null;

    /**
     * With this being non-null the parser will ignore all events more than roughly this many days before now.
     *
     * @var integer
     */
    public $filterDaysAfter = null;

    /**
     * The parsed calendar
     *
     * @var array
     */
    public $cal = array();

    /**
     * Tracks the VFREEBUSY component
     *
     * @var integer
     */
    protected $freeBusyIndex = 0;

    /**
     * Variable to track the previous keyword
     *
     * @var string
     */
    protected $lastKeyword;

    /**
     * Cache valid time zones to avoid unnecessary lookups
     *
     * @var array
     */
    protected $validTimeZones = array();

    /**
     * Event recurrence instances that have been altered
     *
     * @var array
     */
    protected $alteredRecurrenceInstances = array();

    /**
     * An associative array containing ordinal data
     *
     * @var array
     */
    protected $dayOrdinals = array(
        1 => 'first',
        2 => 'second',
        3 => 'third',
        4 => 'fourth',
        5 => 'fifth',
    );

    /**
     * An associative array containing weekday conversion data
     *
     * @var array
     */
    protected $weekdays = array(
        'SU' => 'sunday',
        'MO' => 'monday',
        'TU' => 'tuesday',
        'WE' => 'wednesday',
        'TH' => 'thursday',
        'FR' => 'friday',
        'SA' => 'saturday',
    );

    /**
     * An associative array containing week conversion data
     * (UK = SU, Europe = MO)
     *
     * @var array
     */
    protected $weeks = array(
        'SA' => array('SA', 'SU', 'MO', 'TU', 'WE', 'TH', 'FR'),
        'SU' => array('SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'),
        'MO' => array('MO', 'TU', 'WE', 'TH', 'FR', 'SA', 'SU'),
    );

    /**
     * An associative array containing month names
     *
     * @var array
     */
    protected $monthNames = array(
        1  => 'January',
        2  => 'February',
        3  => 'March',
        4  => 'April',
        5  => 'May',
        6  => 'June',
        7  => 'July',
        8  => 'August',
        9  => 'September',
        10 => 'October',
        11 => 'November',
        12 => 'December',
    );

    /**
     * An associative array containing frequency conversion terms
     *
     * @var array
     */
    protected $frequencyConversion = array(
        'DAILY'   => 'day',
        'WEEKLY'  => 'week',
        'MONTHLY' => 'month',
        'YEARLY'  => 'year',
    );

    /**
     * Holds the username and password for HTTP basic authentication
     *
     * @var array
     */
    protected $httpBasicAuth = array();

    /**
     * Define which variables can be configured
     *
     * @var array
     */
    private static $configurableOptions = array(
        'defaultSpan',
        'defaultTimeZone',
        'defaultWeekStart',
        'disableCharacterReplacement',
        'filterDaysAfter',
        'filterDaysBefore',
        'replaceWindowsTimeZoneIds',
        'skipRecurrence',
        'useTimeZoneWithRRules',
    );

    /**
     * Maps Windows (non-CLDR) time zone ID to IANA ID. This is pragmatic but not 100% precise as one Windows zone ID
     * maps to multiple IANA IDs (one for each territory). For all practical purposes this should be good enough, though.
     *
     * Source: http://unicode.org/repos/cldr/trunk/common/supplemental/windowsZones.xml
     *
     * @var array
     */
    private static $windowsTimeZonesMap = array(
        'AUS Central Standard Time'       => 'Australia/Darwin',
        'AUS Eastern Standard Time'       => 'Australia/Sydney',
        'Afghanistan Standard Time'       => 'Asia/Kabul',
        'Alaskan Standard Time'           => 'America/Anchorage',
        'Aleutian Standard Time'          => 'America/Adak',
        'Altai Standard Time'             => 'Asia/Barnaul',
        'Arab Standard Time'              => 'Asia/Riyadh',
        'Arabian Standard Time'           => 'Asia/Dubai',
        'Arabic Standard Time'            => 'Asia/Baghdad',
        'Argentina Standard Time'         => 'America/Buenos_Aires',
        'Astrakhan Standard Time'         => 'Europe/Astrakhan',
        'Atlantic Standard Time'          => 'America/Halifax',
        'Aus Central W. Standard Time'    => 'Australia/Eucla',
        'Azerbaijan Standard Time'        => 'Asia/Baku',
        'Azores Standard Time'            => 'Atlantic/Azores',
        'Bahia Standard Time'             => 'America/Bahia',
        'Bangladesh Standard Time'        => 'Asia/Dhaka',
        'Belarus Standard Time'           => 'Europe/Minsk',
        'Bougainville Standard Time'      => 'Pacific/Bougainville',
        'Canada Central Standard Time'    => 'America/Regina',
        'Cape Verde Standard Time'        => 'Atlantic/Cape_Verde',
        'Caucasus Standard Time'          => 'Asia/Yerevan',
        'Cen. Australia Standard Time'    => 'Australia/Adelaide',
        'Central America Standard Time'   => 'America/Guatemala',
        'Central Asia Standard Time'      => 'Asia/Almaty',
        'Central Brazilian Standard Time' => 'America/Cuiaba',
        'Central Europe Standard Time'    => 'Europe/Budapest',
        'Central European Standard Time'  => 'Europe/Warsaw',
        'Central Pacific Standard Time'   => 'Pacific/Guadalcanal',
        'Central Standard Time (Mexico)'  => 'America/Mexico_City',
        'Central Standard Time'           => 'America/Chicago',
        'Chatham Islands Standard Time'   => 'Pacific/Chatham',
        'China Standard Time'             => 'Asia/Shanghai',
        'Cuba Standard Time'              => 'America/Havana',
        'Dateline Standard Time'          => 'Etc/GMT+12',
        'E. Africa Standard Time'         => 'Africa/Nairobi',
        'E. Australia Standard Time'      => 'Australia/Brisbane',
        'E. Europe Standard Time'         => 'Europe/Chisinau',
        'E. South America Standard Time'  => 'America/Sao_Paulo',
        'Easter Island Standard Time'     => 'Pacific/Easter',
        'Eastern Standard Time (Mexico)'  => 'America/Cancun',
        'Eastern Standard Time'           => 'America/New_York',
        'Egypt Standard Time'             => 'Africa/Cairo',
        'Ekaterinburg Standard Time'      => 'Asia/Yekaterinburg',
        'FLE Standard Time'               => 'Europe/Kiev',
        'Fiji Standard Time'              => 'Pacific/Fiji',
        'GMT Standard Time'               => 'Europe/London',
        'GTB Standard Time'               => 'Europe/Bucharest',
        'Georgian Standard Time'          => 'Asia/Tbilisi',
        'Greenland Standard Time'         => 'America/Godthab',
        'Greenwich Standard Time'         => 'Atlantic/Reykjavik',
        'Haiti Standard Time'             => 'America/Port-au-Prince',
        'Hawaiian Standard Time'          => 'Pacific/Honolulu',
        'India Standard Time'             => 'Asia/Calcutta',
        'Iran Standard Time'              => 'Asia/Tehran',
        'Israel Standard Time'            => 'Asia/Jerusalem',
        'Jordan Standard Time'            => 'Asia/Amman',
        'Kaliningrad Standard Time'       => 'Europe/Kaliningrad',
        'Korea Standard Time'             => 'Asia/Seoul',
        'Libya Standard Time'             => 'Africa/Tripoli',
        'Line Islands Standard Time'      => 'Pacific/Kiritimati',
        'Lord Howe Standard Time'         => 'Australia/Lord_Howe',
        'Magadan Standard Time'           => 'Asia/Magadan',
        'Magallanes Standard Time'        => 'America/Punta_Arenas',
        'Marquesas Standard Time'         => 'Pacific/Marquesas',
        'Mauritius Standard Time'         => 'Indian/Mauritius',
        'Middle East Standard Time'       => 'Asia/Beirut',
        'Montevideo Standard Time'        => 'America/Montevideo',
        'Morocco Standard Time'           => 'Africa/Casablanca',
        'Mountain Standard Time (Mexico)' => 'America/Chihuahua',
        'Mountain Standard Time'          => 'America/Denver',
        'Myanmar Standard Time'           => 'Asia/Rangoon',
        'N. Central Asia Standard Time'   => 'Asia/Novosibirsk',
        'Namibia Standard Time'           => 'Africa/Windhoek',
        'Nepal Standard Time'             => 'Asia/Katmandu',
        'New Zealand Standard Time'       => 'Pacific/Auckland',
        'Newfoundland Standard Time'      => 'America/St_Johns',
        'Norfolk Standard Time'           => 'Pacific/Norfolk',
        'North Asia East Standard Time'   => 'Asia/Irkutsk',
        'North Asia Standard Time'        => 'Asia/Krasnoyarsk',
        'North Korea Standard Time'       => 'Asia/Pyongyang',
        'Omsk Standard Time'              => 'Asia/Omsk',
        'Pacific SA Standard Time'        => 'America/Santiago',
        'Pacific Standard Time (Mexico)'  => 'America/Tijuana',
        'Pacific Standard Time'           => 'America/Los_Angeles',
        'Pakistan Standard Time'          => 'Asia/Karachi',
        'Paraguay Standard Time'          => 'America/Asuncion',
        'Romance Standard Time'           => 'Europe/Paris',
        'Russia Time Zone 10'             => 'Asia/Srednekolymsk',
        'Russia Time Zone 11'             => 'Asia/Kamchatka',
        'Russia Time Zone 3'              => 'Europe/Samara',
        'Russian Standard Time'           => 'Europe/Moscow',
        'SA Eastern Standard Time'        => 'America/Cayenne',
        'SA Pacific Standard Time'        => 'America/Bogota',
        'SA Western Standard Time'        => 'America/La_Paz',
        'SE Asia Standard Time'           => 'Asia/Bangkok',
        'Saint Pierre Standard Time'      => 'America/Miquelon',
        'Sakhalin Standard Time'          => 'Asia/Sakhalin',
        'Samoa Standard Time'             => 'Pacific/Apia',
        'Sao Tome Standard Time'          => 'Africa/Sao_Tome',
        'Saratov Standard Time'           => 'Europe/Saratov',
        'Singapore Standard Time'         => 'Asia/Singapore',
        'South Africa Standard Time'      => 'Africa/Johannesburg',
        'Sri Lanka Standard Time'         => 'Asia/Colombo',
        'Sudan Standard Time'             => 'Africa/Tripoli',
        'Syria Standard Time'             => 'Asia/Damascus',
        'Taipei Standard Time'            => 'Asia/Taipei',
        'Tasmania Standard Time'          => 'Australia/Hobart',
        'Tocantins Standard Time'         => 'America/Araguaina',
        'Tokyo Standard Time'             => 'Asia/Tokyo',
        'Tomsk Standard Time'             => 'Asia/Tomsk',
        'Tonga Standard Time'             => 'Pacific/Tongatapu',
        'Transbaikal Standard Time'       => 'Asia/Chita',
        'Turkey Standard Time'            => 'Europe/Istanbul',
        'Turks And Caicos Standard Time'  => 'America/Grand_Turk',
        'US Eastern Standard Time'        => 'America/Indianapolis',
        'US Mountain Standard Time'       => 'America/Phoenix',
        'UTC'                             => 'Etc/GMT',
        'UTC+12'                          => 'Etc/GMT-12',
        'UTC+13'                          => 'Etc/GMT-13',
        'UTC-02'                          => 'Etc/GMT+2',
        'UTC-08'                          => 'Etc/GMT+8',
        'UTC-09'                          => 'Etc/GMT+9',
        'UTC-11'                          => 'Etc/GMT+11',
        'Ulaanbaatar Standard Time'       => 'Asia/Ulaanbaatar',
        'Venezuela Standard Time'         => 'America/Caracas',
        'Vladivostok Standard Time'       => 'Asia/Vladivostok',
        'W. Australia Standard Time'      => 'Australia/Perth',
        'W. Central Africa Standard Time' => 'Africa/Lagos',
        'W. Europe Standard Time'         => 'Europe/Berlin',
        'W. Mongolia Standard Time'       => 'Asia/Hovd',
        'West Asia Standard Time'         => 'Asia/Tashkent',
        'West Bank Standard Time'         => 'Asia/Hebron',
        'West Pacific Standard Time'      => 'Pacific/Port_Moresby',
        'Yakutsk Standard Time'           => 'Asia/Yakutsk',
    );

    /**
     * Store the Windows Time Zone IDs to search and replace
     *
     * @var array
     */
    private $windowsTimeZones;

    /**
     * Store the IANA IDs to be used as a replacement for Windows Time Zone IDs
     *
     * @var array
     */
    private $windowsTimeZonesIana;

    /**
     * Creates the ICal object
     *
     * @param  mixed $files
     * @param  array $options
     * @return void
     */
    public function __construct($files = false, array $options = array())
    {
        ini_set('auto_detect_line_endings', '1');

        foreach ($options as $option => $value) {
            if (in_array($option, self::$configurableOptions)) {
                $this->{$option} = $value;
            }
        }

        // Fallback to use the system default time zone
        if (!isset($this->defaultTimeZone) || !$this->isValidTimeZoneId($this->defaultTimeZone)) {
            $this->defaultTimeZone = date_default_timezone_get();
        }

        $this->windowsTimeZones     = array_keys(self::$windowsTimeZonesMap);
        $this->windowsTimeZonesIana = array_values(self::$windowsTimeZonesMap);

        if ($files !== false) {
            $files = is_array($files) ? $files : array($files);

            foreach ($files as $file) {
                if (!is_array($file) && $this->isFileOrUrl($file)) {
                    $lines = $this->fileOrUrl($file);
                } else {
                    $lines = is_array($file) ? $file : array($file);
                }

                $this->initLines($lines);
            }
        }
    }

    /**
     * Initialises lines from a string
     *
     * @param  string $string
     * @return ICal
     */
    public function initString($string)
    {
        if (empty($this->cal)) {
            $lines = explode(PHP_EOL, $string);

            $this->initLines($lines);
        } else {
            trigger_error('ICal::initString: Calendar already initialised in constructor', E_USER_NOTICE);
        }

        return $this;
    }

    /**
     * Initialises lines from a file
     *
     * @param  string $file
     * @return ICal
     */
    public function initFile($file)
    {
        if (empty($this->cal)) {
            $lines = $this->fileOrUrl($file);

            $this->initLines($lines);
        } else {
            trigger_error('ICal::initFile: Calendar already initialised in constructor', E_USER_NOTICE);
        }

        return $this;
    }

    /**
     * Initialises lines from a URL
     *
     * @param  string $url
     * @param  string $username
     * @param  string $password
     * @return ICal
     */
    public function initUrl($url, $username = null, $password = null)
    {
        if (!is_null($username) && !is_null($password)) {
            $this->httpBasicAuth['username'] = $username;
            $this->httpBasicAuth['password'] = $password;
        }

        $this->initFile($url);

        return $this;
    }

    /**
     * Initialises the parser using an array
     * containing each line of iCal content
     *
     * @param  array $lines
     * @return void
     */
    protected function initLines(array $lines)
    {
        $lines = $this->unfold($lines);

        if (stristr($lines[0], 'BEGIN:VCALENDAR') !== false) {
            $component = '';
            foreach ($lines as $line) {
                $line = rtrim($line); // Trim trailing whitespace
                $line = $this->removeUnprintableChars($line);

                if (!$this->disableCharacterReplacement) {
                    $line = $this->cleanData($line);
                }

                if ($this->replaceWindowsTimeZoneIds && strpos($line, 'TZID') !== false) {
                    $line = $this->replaceWindowsTimeZoneId($line);
                }

                $add = $this->keyValueFromString($line);

                $keyword = $add[0];
                $values  = $add[1]; // May be an array containing multiple values

                if (!is_array($values)) {
                    if (!empty($values)) {
                        $values = array($values); // Make an array as not already
                        $blankArray = array(); // Empty placeholder array
                        array_push($values, $blankArray);
                    } else {
                        $values = array(); // Use blank array to ignore this line
                    }
                } elseif (empty($values[0])) {
                    $values = array(); // Use blank array to ignore this line
                }

                // Reverse so that our array of properties is processed first
                $values = array_reverse($values);

                foreach ($values as $value) {
                    switch ($line) {
                        // https://www.kanzaki.com/docs/ical/vtodo.html
                        case 'BEGIN:VTODO':
                            if (!is_array($value)) {
                                $this->todoCount++;
                            }

                            $component = 'VTODO';
                        break;

                        // https://www.kanzaki.com/docs/ical/vevent.html
                        case 'BEGIN:VEVENT':
                            if (!is_array($value)) {
                                $this->eventCount++;
                            }

                            $component = 'VEVENT';
                        break;

                        // https://www.kanzaki.com/docs/ical/vfreebusy.html
                        case 'BEGIN:VFREEBUSY':
                            if (!is_array($value)) {
                                $this->freeBusyIndex++;
                            }

                            $component = 'VFREEBUSY';
                        break;

                        case 'BEGIN:VALARM':
                            if (!is_array($value)) {
                                $this->alarmCount++;
                            }

                            $component = 'VALARM';
                        break;

                        case 'END:VALARM':
                            $component = 'VEVENT';
                        break;

                        case 'BEGIN:DAYLIGHT':
                        case 'BEGIN:STANDARD':
                        case 'BEGIN:VCALENDAR':
                        case 'BEGIN:VTIMEZONE':
                            $component = $value;
                        break;

                        case 'END:DAYLIGHT':
                        case 'END:STANDARD':
                        case 'END:VCALENDAR':
                        case 'END:VEVENT':
                        case 'END:VFREEBUSY':
                        case 'END:VTIMEZONE':
                        case 'END:VTODO':
                            $component = 'VCALENDAR';
                        break;

                        default:
                            $this->addCalendarComponentWithKeyAndValue($component, $keyword, $value);
                        break;
                    }
                }
            }

            $this->processEvents();

            if (!$this->skipRecurrence) {
                $this->processRecurrences();

                // Apply changes to altered recurrence instances
                if (!empty($this->alteredRecurrenceInstances)) {
                    $events = $this->cal['VEVENT'];

                    foreach ($this->alteredRecurrenceInstances as $alteredRecurrenceInstance) {
                        if (isset($alteredRecurrenceInstance['altered-event'])) {
                            $alteredEvent = $alteredRecurrenceInstance['altered-event'];
                            $key          = key($alteredEvent);
                            $events[$key] = $alteredEvent[$key];
                        }
                    }

                    $this->cal['VEVENT'] = $events;
                }
            }

            if (!is_null($this->filterDaysBefore) || !is_null($this->filterDaysAfter)) {
                $this->reduceEventsToMinMaxRange();
            }

            $this->processDateConversions();
        }
    }

    /**
     * Reduces the number of events to the defined minimum and maximum range
     *
     * @return void
     */
    protected function reduceEventsToMinMaxRange()
    {
        $events = (isset($this->cal['VEVENT'])) ? $this->cal['VEVENT'] : array();

        if (!empty($events)) {
            // Ideally you would use `PHP_INT_MIN` from PHP 7
            $php_int_min = -2147483648;

            $minTimestamp = is_null($this->filterDaysBefore) ? $php_int_min : (new \DateTime('now'))->sub(new \DateInterval('P' . $this->filterDaysBefore . 'D'))->getTimestamp();
            $maxTimestamp = is_null($this->filterDaysAfter) ? PHP_INT_MAX : (new \DateTime('now'))->add(new \DateInterval('P' . $this->filterDaysAfter . 'D'))->getTimestamp();

            foreach ($events as $key => $anEvent) {
                if (!$this->isValidDate($anEvent['DTSTART']) || $this->isOutOfRange($anEvent['DTSTART'], $minTimestamp, $maxTimestamp)) {
                    $this->eventCount--;

                    unset($events[$key]);

                    continue;
                }
            }

            $this->cal['VEVENT'] = $events;
        }
    }

    /**
     * Determines whether an event's start time is within a given range
     *
     * @param  string  $eventStart
     * @param  integer $minTimestamp
     * @param  integer $maxTimestamp
     * @return boolean
     */
    protected function isOutOfRange($eventStart, $minTimestamp, $maxTimestamp)
    {
        $eventStartTimestamp = strtotime(explode('T', $eventStart)[0]);

        return $eventStartTimestamp < $minTimestamp || $eventStartTimestamp > $maxTimestamp;
    }

    /**
     * Unfolds an iCal file in preparation for parsing
     * (https://icalendar.org/iCalendar-RFC-5545/3-1-content-lines.html)
     *
     * @param  array $lines
     * @return array
     */
    protected function unfold(array $lines)
    {
        $string = implode(PHP_EOL, $lines);
        $string = preg_replace('/' . PHP_EOL . '[ \t]/', '', $string);
        $lines  = explode(PHP_EOL, $string);

        return $lines;
    }

    /**
     * Add one key and value pair to the `$this->cal` array
     *
     * @param  string         $component
     * @param  string|boolean $keyword
     * @param  string         $value
     * @return void
     */
    protected function addCalendarComponentWithKeyAndValue($component, $keyword, $value)
    {
        if ($keyword == false) {
            $keyword = $this->lastKeyword;
        }

        switch ($component) {
            case 'VALARM':
                $key1 = 'VEVENT';
                $key2 = ($this->eventCount - 1);
                $key3 = $component;

                if (!isset($this->cal[$key1][$key2][$key3]["{$keyword}_array"])) {
                    $this->cal[$key1][$key2][$key3]["{$keyword}_array"] = array();
                }

                if (is_array($value)) {
                    // Add array of properties to the end
                    array_push($this->cal[$key1][$key2][$key3]["{$keyword}_array"], $value);
                } else {
                    if (!isset($this->cal[$key1][$key2][$key3][$keyword])) {
                        $this->cal[$key1][$key2][$key3][$keyword] = $value;
                    }

                    if ($this->cal[$key1][$key2][$key3][$keyword] !== $value) {
                        $this->cal[$key1][$key2][$key3][$keyword] .= ',' . $value;
                    }
                }
            break;

            case 'VEVENT':
                $key1 = $component;
                $key2 = ($this->eventCount - 1);

                if (!isset($this->cal[$key1][$key2]["{$keyword}_array"])) {
                    $this->cal[$key1][$key2]["{$keyword}_array"] = array();
                }

                if (is_array($value)) {
                    // Add array of properties to the end
                    array_push($this->cal[$key1][$key2]["{$keyword}_array"], $value);
                } else {
                    if (!isset($this->cal[$key1][$key2][$keyword])) {
                        $this->cal[$key1][$key2][$keyword] = $value;
                    }

                    if ($keyword === 'EXDATE') {
                        if (trim($value) === $value) {
                            $array = array_filter(explode(',', $value));
                            $this->cal[$key1][$key2]["{$keyword}_array"][] = $array;
                        } else {
                            $value = explode(',', implode(',', $this->cal[$key1][$key2]["{$keyword}_array"][1]) . trim($value));
                            $this->cal[$key1][$key2]["{$keyword}_array"][1] = $value;
                        }
                    } else {
                        $this->cal[$key1][$key2]["{$keyword}_array"][] = $value;

                        if ($keyword === 'DURATION') {
                            $duration = new \DateInterval($value);
                            array_push($this->cal[$key1][$key2]["{$keyword}_array"], $duration);
                        }
                    }

                    if ($this->cal[$key1][$key2][$keyword] !== $value) {
                        $this->cal[$key1][$key2][$keyword] .= ',' . $value;
                    }
                }
            break;

            case 'VFREEBUSY':
                $key1 = $component;
                $key2 = ($this->freeBusyIndex - 1);
                $key3 = $keyword;

                if ($keyword === 'FREEBUSY') {
                    if (is_array($value)) {
                        $this->cal[$key1][$key2][$key3][][] = $value;
                    } else {
                        $this->freeBusyCount++;

                        end($this->cal[$key1][$key2][$key3]);
                        $key = key($this->cal[$key1][$key2][$key3]);

                        $value = explode('/', $value);
                        $this->cal[$key1][$key2][$key3][$key][] = $value;
                    }
                } else {
                    $this->cal[$key1][$key2][$key3][] = $value;
                }
            break;

            case 'VTODO':
                $this->cal[$component][$this->todoCount - 1][$keyword] = $value;
            break;

            default:
                $this->cal[$component][$keyword] = $value;
            break;
        }

        $this->lastKeyword = $keyword;
    }

    /**
     * Gets the key value pair from an iCal string
     *
     * @param  string $text
     * @return array|boolean
     */
    protected function keyValueFromString($text)
    {
        $text = htmlspecialchars($text, ENT_NOQUOTES, 'UTF-8');

        $colon = strpos($text, ':');
        $quote = strpos($text, '"');
        if ($colon === false) {
            $matches = array();
        } elseif ($quote === false || $colon < $quote) {
            list($before, $after) = explode(':', $text, 2);
            $matches              = array($text, $before, $after);
        } else {
            list($before, $text) = explode('"', $text, 2);
            $text                = '"' . $text;
            $matches             = str_getcsv($text, ':');
            $combinedValue       = '';

            foreach ($matches as $key => $match) {
                if ($key === 0) {
                    if (!empty($before)) {
                        $matches[$key] = $before . '"' . $matches[$key] . '"';
                    }
                } else {
                    if ($key > 1) {
                        $combinedValue .= ':';
                    }

                    $combinedValue .= $matches[$key];
                }
            }

            $matches    = array_slice($matches, 0, 2);
            $matches[1] = $combinedValue;
            array_unshift($matches, $before . $text);
        }

        if (count($matches) === 0) {
            return false;
        }

        if (preg_match('/^([A-Z-]+)([;][\w\W]*)?$/', $matches[1])) {
            $matches = array_splice($matches, 1, 2); // Remove first match and re-align ordering

            // Process properties
            if (preg_match('/([A-Z-]+)[;]([\w\W]*)/', $matches[0], $properties)) {
                // Remove first match
                array_shift($properties);
                // Fix to ignore everything in keyword after a ; (e.g. Language, TZID, etc.)
                $matches[0] = $properties[0];
                array_shift($properties); // Repeat removing first match

                $formatted = array();
                foreach ($properties as $property) {
                    // Match semicolon separator outside of quoted substrings
                    preg_match_all('~[^' . PHP_EOL . '";]+(?:"[^"\\\]*(?:\\\.[^"\\\]*)*"[^' . PHP_EOL . '";]*)*~', $property, $attributes);
                    // Remove multi-dimensional array and use the first key
                    $attributes = (sizeof($attributes) === 0) ? array($property) : reset($attributes);

                    if (is_array($attributes)) {
                        foreach ($attributes as $attribute) {
                            // Match equals sign separator outside of quoted substrings
                            preg_match_all(
                                '~[^' . PHP_EOL . '"=]+(?:"[^"\\\]*(?:\\\.[^"\\\]*)*"[^' . PHP_EOL . '"=]*)*~',
                                $attribute,
                                $values
                            );
                            // Remove multi-dimensional array and use the first key
                            $value = (sizeof($values) === 0) ? null : reset($values);

                            if (is_array($value) && isset($value[1])) {
                                // Remove double quotes from beginning and end only
                                $formatted[$value[0]] = trim($value[1], '"');
                            }
                        }
                    }
                }

                // Assign the keyword property information
                $properties[0] = $formatted;

                // Add match to beginning of array
                array_unshift($properties, $matches[1]);
                $matches[1] = $properties;
            }

            return $matches;
        } else {
            return false; // Ignore this match
        }
    }

    /**
     * Returns a `DateTime` object from an iCal date time format
     *
     * @param  string  $icalDate
     * @param  boolean $forceTimeZone
     * @param  boolean $forceUtc
     * @return \DateTime
     * @throws \Exception
     */
    public function iCalDateToDateTime($icalDate, $forceTimeZone = false, $forceUtc = false)
    {
        /**
         * iCal times may be in 3 formats, (https://www.kanzaki.com/docs/ical/dateTime.html)
         *
         * UTC:      Has a trailing 'Z'
         * Floating: No time zone reference specified, no trailing 'Z', use local time
         * TZID:     Set time zone as specified
         *
         * Use DateTime class objects to get around limitations with `mktime` and `gmmktime`.
         * Must have a local time zone set to process floating times.
         */
        $pattern  = '/\AT?Z?I?D?=?(.*):?'; // [1]: Time zone
        $pattern .= '([0-9]{4})';          // [2]: YYYY
        $pattern .= '([0-9]{2})';          // [3]: MM
        $pattern .= '([0-9]{2})';          // [4]: DD
        $pattern .= 'T?';                  //      Time delimiter
        $pattern .= '([0-9]{0,2})';        // [5]: HH
        $pattern .= '([0-9]{0,2})';        // [6]: MM
        $pattern .= '([0-9]{0,2})';        // [7]: SS
        $pattern .= '(Z?)/';               // [8]: UTC flag

        preg_match($pattern, $icalDate, $date);

        if (empty($date)) {
            throw new \Exception('Invalid iCal date format.');
        }

        // A Unix timestamp cannot represent a date prior to 1 Jan 1970
        $year  = $date[2];
        $isUtc = false;

        if ($year <= self::UNIX_MIN_YEAR) {
            $eventTimeZone = ltrim(strstr($icalDate, ':', true), 'TZID=');

            if (empty($eventTimeZone)) {
                $dateTime = new \DateTime($icalDate, new \DateTimeZone($this->defaultTimeZone));
            } else {
                $icalDate = ltrim(strstr($icalDate, ':'), ':');
                $dateTime = new \DateTime($icalDate, new \DateTimeZone($eventTimeZone));
            }
        } else {
            if ($forceTimeZone) {
                // TZID={Time Zone}:
                if (isset($date[1])) {
                    $eventTimeZone = rtrim($date[1], ':');
                }

                if ($date[8] === 'Z') {
                    $isUtc    = true;
                    $dateTime = new \DateTime('now', new \DateTimeZone(self::TIME_ZONE_UTC));
                } elseif (isset($eventTimeZone) && $this->isValidIanaTimeZoneId($eventTimeZone)) {
                    $dateTime = new \DateTime('now', new \DateTimeZone($eventTimeZone));
                } elseif (isset($eventTimeZone) && $this->isValidCldrTimeZoneId($eventTimeZone)) {
                    $dateTime = new \DateTime('now', new \DateTimeZone($this->isValidCldrTimeZoneId($eventTimeZone, true)));
                } else {
                    $dateTime = new \DateTime('now', new \DateTimeZone($this->defaultTimeZone));
                }
            } else {
                if ($forceUtc) {
                    $dateTime = new \DateTime('now', new \DateTimeZone(self::TIME_ZONE_UTC));
                } else {
                    $dateTime = new \DateTime('now');
                }
            }

            $dateTime->setDate((int) $date[2], (int) $date[3], (int) $date[4]);
            $dateTime->setTime((int) $date[5], (int) $date[6], (int) $date[7]);
        }

        if ($forceTimeZone && $isUtc) {
            $dateTime->setTimezone(new \DateTimeZone($this->defaultTimeZone));
        } elseif ($forceUtc) {
            $dateTime->setTimezone(new \DateTimeZone(self::TIME_ZONE_UTC));
        }

        return $dateTime;
    }

    /**
     * Returns a Unix timestamp from an iCal date time format
     *
     * @param  string  $icalDate
     * @param  boolean $forceTimeZone
     * @param  boolean $forceUtc
     * @return integer
     */
    public function iCalDateToUnixTimestamp($icalDate, $forceTimeZone = false, $forceUtc = false)
    {
        $dateTime = $this->iCalDateToDateTime($icalDate, $forceTimeZone, $forceUtc);
        $offset   = 0;

        if ($forceTimeZone) {
            $offset = $dateTime->getOffset();
        }

        return $dateTime->getTimestamp() + $offset;
    }

    /**
     * Returns a date adapted to the calendar time zone depending on the event `TZID`
     *
     * @param  array  $event
     * @param  string $key
     * @param  string $format
     * @return string|boolean
     */
    public function iCalDateWithTimeZone(array $event, $key, $format = self::DATE_TIME_FORMAT)
    {
        if (!isset($event[$key . '_array']) || !isset($event[$key])) {
            return false;
        }

        $dateArray = $event[$key . '_array'];

        if ($key === 'DURATION') {
            $duration = end($dateArray);
            $dateTime = $this->parseDuration($event['DTSTART'], $duration, null);
        } else {
            $dateTime = new \DateTime($dateArray[1], new \DateTimeZone(self::TIME_ZONE_UTC));
            $dateTime->setTimezone(new \DateTimeZone($this->calendarTimeZone()));
        }

        // Force time zone
        if (isset($dateArray[0]['TZID'])) {
            if ($this->isValidIanaTimeZoneId($dateArray[0]['TZID'])) {
                $dateTime->setTimezone(new \DateTimeZone($dateArray[0]['TZID']));
            } elseif ($this->isValidCldrTimeZoneId($dateArray[0]['TZID'])) {
                $dateTime->setTimezone(new \DateTimeZone($this->isValidCldrTimeZoneId($dateArray[0]['TZID'], true)));
            } else {
                $dateTime->setTimezone(new \DateTimeZone($this->defaultTimeZone));
            }
        }

        if (is_null($format)) {
            $output = $dateTime;
        } else {
            if ($format === self::UNIX_FORMAT) {
                $output = $dateTime->getTimestamp();
            } else {
                $output = $dateTime->format($format);
            }
        }

        return $output;
    }

    /**
     * Performs admin tasks on all events as read from the iCal file.
     * Adds a Unix timestamp to all `{DTSTART|DTEND|RECURRENCE-ID}_array` arrays
     * Tracks modified recurrence instances
     *
     * @return void
     */
    protected function processEvents()
    {
        $events = (isset($this->cal['VEVENT'])) ? $this->cal['VEVENT'] : array();

        if (!empty($events)) {
            foreach ($events as $key => $anEvent) {
                foreach (array('DTSTART', 'DTEND', 'RECURRENCE-ID') as $type) {
                    if (isset($anEvent[$type])) {
                        $date = $anEvent[$type . '_array'][1];

                        if (isset($anEvent[$type . '_array'][0]['TZID'])) {
                            $date = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent[$type . '_array'][0]['TZID']) . $date;
                        }

                        $anEvent[$type . '_array'][2] = $this->iCalDateToUnixTimestamp($date, true, true);
                        $anEvent[$type . '_array'][3] = $date;
                    }
                }

                if (isset($anEvent['RECURRENCE-ID'])) {
                    $uid = $anEvent['UID'];

                    if (!isset($this->alteredRecurrenceInstances[$uid])) {
                        $this->alteredRecurrenceInstances[$uid] = array();
                    }

                    $recurrenceDateUtc = $this->iCalDateToUnixTimestamp($anEvent['RECURRENCE-ID_array'][3], true, true);
                    $this->alteredRecurrenceInstances[$uid][$key] = $recurrenceDateUtc;
                }

                $events[$key] = $anEvent;
            }

            $eventKeysToRemove = array();

            foreach ($events as $key => $event) {
                $checks[] = !isset($event['RECURRENCE-ID']);
                $checks[] = isset($event['UID']);
                $checks[] = isset($event['UID']) && isset($this->alteredRecurrenceInstances[$event['UID']]);

                if ((bool) array_product($checks)) {
                    $eventDtstartUnix = $this->iCalDateToUnixTimestamp($event['DTSTART_array'][3], true, true);

                    if (false !== $alteredEventKey = array_search($eventDtstartUnix, $this->alteredRecurrenceInstances[$event['UID']])) {
                        $eventKeysToRemove[] = $alteredEventKey;

                        $alteredEvent = array_replace_recursive($events[$key], $events[$alteredEventKey]);
                        $this->alteredRecurrenceInstances[$event['UID']]['altered-event'] = array($key => $alteredEvent);
                    }
                }

                unset($checks);
            }

            if (!empty($eventKeysToRemove)) {
                foreach ($eventKeysToRemove as $eventKeyToRemove) {
                    $events[$eventKeyToRemove] = null;
                }
            }

            $this->cal['VEVENT'] = $events;
        }
    }

    /**
     * Processes recurrence rules
     *
     * @return void
     */
    protected function processRecurrences()
    {
        $events = (isset($this->cal['VEVENT'])) ? $this->cal['VEVENT'] : array();

        $recurrenceEvents    = array();
        $allRecurrenceEvents = array();

        if (!empty($events)) {
            foreach ($events as $anEvent) {
                if (isset($anEvent['RRULE']) && $anEvent['RRULE'] !== '') {
                    // Tag as generated by a recurrence rule
                    $anEvent['RRULE_array'][2] = self::RECURRENCE_EVENT;

                    $countNb = 0;

                    $isAllDayEvent = (strlen($anEvent['DTSTART_array'][1]) === 8) ? true : false;

                    $initialStart             = new \DateTime($anEvent['DTSTART_array'][1]);
                    $initialStartTimeZoneName = $initialStart->getTimezone()->getName();

                    if (isset($anEvent['DTEND'])) {
                        $initialEnd             = new \DateTime($anEvent['DTEND_array'][1]);
                        $initialEndTimeZoneName = $initialEnd->getTimezone()->getName();
                    } else {
                        $initialEndTimeZoneName = $initialStartTimeZoneName;
                    }

                    // Recurring event, parse RRULE and add appropriate duplicate events
                    $rrules = array();
                    $rruleStrings = explode(';', $anEvent['RRULE']);

                    foreach ($rruleStrings as $s) {
                        list($k, $v) = explode('=', $s);
                        $rrules[$k] = $v;
                    }

                    // Get frequency
                    $frequency = $rrules['FREQ'];
                    // Get Start timestamp
                    $startTimestamp = $initialStart->getTimestamp();

                    if (isset($anEvent['DTEND'])) {
                        $endTimestamp = $initialEnd->getTimestamp();
                    } elseif (isset($anEvent['DURATION'])) {
                        $duration = end($anEvent['DURATION_array']);
                        $endTimestamp = $this->parseDuration($anEvent['DTSTART'], $duration);
                    } else {
                        $endTimestamp = $anEvent['DTSTART_array'][2];
                    }

                    $eventTimestampOffset = $endTimestamp - $startTimestamp;
                    // Get Interval
                    $interval = (isset($rrules['INTERVAL']) && $rrules['INTERVAL'] !== '') ? $rrules['INTERVAL'] : 1;

                    $dayNumber = null;
                    $weekday   = null;

                    if (in_array($frequency, array('MONTHLY', 'YEARLY')) && isset($rrules['BYDAY']) && $rrules['BYDAY'] !== '') {
                        // Deal with BYDAY
                        $byDay     = $rrules['BYDAY'];
                        $dayNumber = intval($byDay);

                        if (empty($dayNumber)) { // Returns 0 when no number defined in BYDAY
                            if (!isset($rrules['BYSETPOS'])) {
                                $dayNumber = 1; // Set first as default
                            } elseif (is_numeric($rrules['BYSETPOS'])) {
                                $dayNumber = $rrules['BYSETPOS'];
                            }
                        }

                        $weekday = substr($byDay, -2);
                    }

                    if (is_int($this->defaultSpan)) {
                        $untilDefault = date_create('now');
                        $untilDefault->modify($this->defaultSpan . ' year');
                        $untilDefault->setTime(23, 59, 59); // End of the day
                    } else {
                        trigger_error('ICal::defaultSpan: User defined value is not an integer', E_USER_NOTICE);
                    }

                    // Compute EXDATEs
                    $exdates = $this->parseExdates($anEvent);

                    $countOrig = null;

                    if (isset($rrules['UNTIL'])) {
                        // Get Until
                        $until = strtotime($rrules['UNTIL']);
                    } elseif (isset($rrules['COUNT'])) {
                        $countOrig = (is_numeric($rrules['COUNT']) && $rrules['COUNT'] > 1) ? $rrules['COUNT'] : 0;

                        // Increment count by the number of excluded dates
                        $countOrig += sizeof($exdates);

                        // Remove one to exclude the occurrence that initialises the rule
                        $count = ($countOrig - 1);

                        if ($interval >= 2) {
                            $count += ($count > 0) ? ($count * $interval) : 0;
                        }

                        $countNb = 1;
                        $offset  = "+{$count} " . $this->frequencyConversion[$frequency];
                        $until   = strtotime($offset, $startTimestamp);

                        if (in_array($frequency, array('MONTHLY', 'YEARLY'))
                            && isset($rrules['BYDAY']) && $rrules['BYDAY'] !== ''
                        ) {
                            $dtstart = date_create($anEvent['DTSTART']);

                            if (!$dtstart) {
                                continue;
                            }

                            for ($i = 1; $i <= $count; $i++) {
                                $dtstartClone = clone $dtstart;
                                $dtstartClone->modify('next ' . $this->frequencyConversion[$frequency]);
                                $offset = "{$this->convertDayOrdinalToPositive($dayNumber, $weekday, $dtstartClone)} {$this->weekdays[$weekday]} of " . $dtstartClone->format('F Y H:i:01');
                                $dtstart->modify($offset);
                            }

                            // Jumping X months forwards doesn't mean
                            // the end date will fall on the same day defined in BYDAY
                            // Use the largest of these to ensure we are going far enough
                            // in the future to capture our final end day
                            $until = max($until, $dtstart->format(self::UNIX_FORMAT));
                        }

                        unset($offset);
                    } elseif (isset($untilDefault)) {
                        $until = $untilDefault->getTimestamp();
                    }

                    $until = intval($until);

                    // Decide how often to add events and do so
                    switch ($frequency) {
                        case 'DAILY':
                            // Simply add a new event each interval of days until UNTIL is reached
                            $offset = "+{$interval} day";
                            $recurringTimestamp = strtotime($offset, $startTimestamp);

                            while ($recurringTimestamp <= $until) {
                                $dayRecurringTimestamp = $recurringTimestamp;

                                // Adjust time zone from initial event
                                $dayRecurringOffset = 0;
                                if ($this->useTimeZoneWithRRules) {
                                    $recurringTimeZone = \DateTime::createFromFormat(self::UNIX_FORMAT, $dayRecurringTimestamp);
                                    $recurringTimeZone->setTimezone($initialStart->getTimezone());
                                    $dayRecurringOffset = $recurringTimeZone->getOffset();
                                    $dayRecurringTimestamp += $dayRecurringOffset;
                                }

                                // Add event
                                $anEvent['DTSTART'] = date(self::DATE_TIME_FORMAT, $dayRecurringTimestamp) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');
                                $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                $anEvent['DTSTART_array'][2] = $dayRecurringTimestamp;
                                $anEvent['DTEND_array']      = $anEvent['DTSTART_array'];
                                $anEvent['DTEND_array'][2]  += $eventTimestampOffset;
                                $anEvent['DTEND'] = date(
                                        self::DATE_TIME_FORMAT,
                                        $anEvent['DTEND_array'][2]
                                    ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');
                                $anEvent['DTEND_array'][1] = $anEvent['DTEND'];

                                // Exclusions
                                $isExcluded = array_filter($exdates, function ($exdate) use ($anEvent, $dayRecurringOffset) {
                                    return self::isExdateMatch($exdate, $anEvent, $dayRecurringOffset);
                                });

                                if (isset($anEvent['UID'])) {
                                    $searchDate = $anEvent['DTSTART'];
                                    if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                                        $searchDate = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent['DTSTART_array'][0]['TZID']) . $searchDate;
                                    }

                                    if (isset($this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                        $searchDateUtc = $this->iCalDateToUnixTimestamp($searchDate, true, true);
                                        if (in_array($searchDateUtc, $this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                            $isExcluded = true;
                                        }
                                    }
                                }

                                if (!$isExcluded) {
                                    $anEvent            = $this->processEventIcalDateTime($anEvent);
                                    $recurrenceEvents[] = $anEvent;
                                    $this->eventCount++;

                                    // If RRULE[COUNT] is reached then break
                                    if (isset($rrules['COUNT'])) {
                                        $countNb++;

                                        if ($countNb >= $countOrig) {
                                            break;
                                        }
                                    }
                                }

                                // Move forwards
                                $recurringTimestamp = strtotime($offset, $recurringTimestamp);
                            }

                            $recurrenceEvents    = $this->trimToRecurrenceCount($rrules, $recurrenceEvents);
                            $allRecurrenceEvents = array_merge($allRecurrenceEvents, $recurrenceEvents);
                            $recurrenceEvents    = array(); // Reset
                        break;

                        case 'WEEKLY':
                            // Create offset
                            $offset = "+{$interval} week";

                            $wkst  = (isset($rrules['WKST']) && in_array($rrules['WKST'], array('SA', 'SU', 'MO'))) ? $rrules['WKST'] : $this->defaultWeekStart;
                            $aWeek = $this->weeks[$wkst];
                            $days  = array('SA' => 'Saturday', 'SU' => 'Sunday', 'MO' => 'Monday');

                            // Build list of days of week to add events
                            $weekdays = $aWeek;

                            if (isset($rrules['BYDAY']) && $rrules['BYDAY'] !== '') {
                                $byDays = explode(',', $rrules['BYDAY']);
                            } else {
                                // A textual representation of a day, two letters (e.g. SU)
                                $byDays = array(mb_substr(strtoupper($initialStart->format('D')), 0, 2));
                            }

                            // Get timestamp of first day of start week
                            $weekRecurringTimestamp = (strcasecmp($initialStart->format('l'), $this->weekdays[$wkst]) === 0)
                                ? $startTimestamp
                                : strtotime("last {$days[$wkst]} " . $initialStart->format('H:i:s'), $startTimestamp);

                            // Step through weeks
                            while ($weekRecurringTimestamp <= $until) {
                                $dayRecurringTimestamp = $weekRecurringTimestamp;

                                // Adjust time zone from initial event
                                $dayRecurringOffset = 0;
                                if ($this->useTimeZoneWithRRules) {
                                    $dayRecurringTimeZone = \DateTime::createFromFormat(self::UNIX_FORMAT, $dayRecurringTimestamp);
                                    $dayRecurringTimeZone->setTimezone($initialStart->getTimezone());
                                    $dayRecurringOffset = $dayRecurringTimeZone->getOffset();
                                    $dayRecurringTimestamp += $dayRecurringOffset;
                                }

                                foreach ($weekdays as $day) {
                                    // Check if day should be added
                                    if (in_array($day, $byDays) && $dayRecurringTimestamp > $startTimestamp
                                        && $dayRecurringTimestamp <= $until
                                    ) {
                                        // Add event
                                        $anEvent['DTSTART'] = date(self::DATE_TIME_FORMAT, $dayRecurringTimestamp) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');
                                        $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                        $anEvent['DTSTART_array'][2] = $dayRecurringTimestamp;
                                        $anEvent['DTEND_array']      = $anEvent['DTSTART_array'];
                                        $anEvent['DTEND_array'][2]  += $eventTimestampOffset;
                                        $anEvent['DTEND'] = date(
                                                self::DATE_TIME_FORMAT,
                                                $anEvent['DTEND_array'][2]
                                            ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');
                                        $anEvent['DTEND_array'][1] = $anEvent['DTEND'];

                                        // Exclusions
                                        $isExcluded = array_filter($exdates, function ($exdate) use ($anEvent, $dayRecurringOffset) {
                                            return self::isExdateMatch($exdate, $anEvent, $dayRecurringOffset);
                                        });

                                        if (isset($anEvent['UID'])) {
                                            $searchDate = $anEvent['DTSTART'];
                                            if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                                                $searchDate = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent['DTSTART_array'][0]['TZID']) . $searchDate;
                                            }

                                            if (isset($this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                $searchDateUtc = $this->iCalDateToUnixTimestamp($searchDate, true, true);
                                                if (in_array($searchDateUtc, $this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                    $isExcluded = true;
                                                }
                                            }
                                        }

                                        if (!$isExcluded) {
                                            $anEvent            = $this->processEventIcalDateTime($anEvent);
                                            $recurrenceEvents[] = $anEvent;
                                            $this->eventCount++;

                                            // If RRULE[COUNT] is reached then break
                                            if (isset($rrules['COUNT'])) {
                                                $countNb++;

                                                if ($countNb >= $countOrig) {
                                                    break 2;
                                                }
                                            }
                                        }
                                    }

                                    // Move forwards a day
                                    $dayRecurringTimestamp = strtotime('+1 day', $dayRecurringTimestamp);
                                }

                                // Move forwards $interval weeks
                                $weekRecurringTimestamp = strtotime($offset, $weekRecurringTimestamp);
                            }

                            $recurrenceEvents    = $this->trimToRecurrenceCount($rrules, $recurrenceEvents);
                            $allRecurrenceEvents = array_merge($allRecurrenceEvents, $recurrenceEvents);
                            $recurrenceEvents    = array(); // Reset
                        break;

                        case 'MONTHLY':
                            // Create offset
                            $recurringTimestamp = $startTimestamp;
                            $offset = "+{$interval} month";

                            if (isset($rrules['BYMONTHDAY']) && $rrules['BYMONTHDAY'] !== '') {
                                // Deal with BYMONTHDAY
                                $monthdays = explode(',', $rrules['BYMONTHDAY']);

                                while ($recurringTimestamp <= $until) {
                                    foreach ($monthdays as $key => $monthday) {
                                        $monthRecurringTimestamp = null;

                                        if ($key === 0) {
                                            // Ensure original event conforms to monthday rule
                                            $anEvent['DTSTART'] = gmdate(
                                                    'Ym' . sprintf('%02d', $monthday) . '\T' . self::TIME_FORMAT,
                                                    strtotime($anEvent['DTSTART'])
                                                ) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');

                                            $anEvent['DTEND'] = gmdate(
                                                    'Ym' . sprintf('%02d', $monthday) . '\T' . self::TIME_FORMAT,
                                                    isset($anEvent['DURATION'])
                                                        ? $this->parseDuration($anEvent['DTSTART'], end($anEvent['DURATION_array']))
                                                        : strtotime($anEvent['DTEND'])
                                                ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');

                                            $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                            $anEvent['DTSTART_array'][2] = $this->iCalDateToUnixTimestamp($anEvent['DTSTART']);
                                            $anEvent['DTEND_array'][1]   = $anEvent['DTEND'];
                                            $anEvent['DTEND_array'][2]   = $this->iCalDateToUnixTimestamp($anEvent['DTEND']);

                                            // Ensure recurring timestamp confirms to BYMONTHDAY rule
                                            $monthRecurringTimestamp = $this->iCalDateToUnixTimestamp(
                                                gmdate(
                                                    'Ym' . sprintf('%02d', $monthday) . '\T' . self::TIME_FORMAT,
                                                    $recurringTimestamp
                                                ) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '')
                                            );
                                        }

                                        // Adjust time zone from initial event
                                        $monthRecurringOffset = 0;
                                        if ($this->useTimeZoneWithRRules) {
                                            $recurringTimeZone = \DateTime::createFromFormat(self::UNIX_FORMAT, $monthRecurringTimestamp);
                                            $recurringTimeZone->setTimezone($initialStart->getTimezone());
                                            $monthRecurringOffset = $recurringTimeZone->getOffset();
                                            $monthRecurringTimestamp += $monthRecurringOffset;
                                        }

                                        // Add event
                                        $anEvent['DTSTART'] = date(
                                                'Ym' . sprintf('%02d', $monthday) . '\T' . self::TIME_FORMAT,
                                                $monthRecurringTimestamp
                                            ) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');
                                        $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                        $anEvent['DTSTART_array'][2] = $monthRecurringTimestamp;
                                        $anEvent['DTEND_array']      = $anEvent['DTSTART_array'];
                                        $anEvent['DTEND_array'][2]  += $eventTimestampOffset;
                                        $anEvent['DTEND'] = date(
                                                self::DATE_TIME_FORMAT,
                                                $anEvent['DTEND_array'][2]
                                            ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');
                                        $anEvent['DTEND_array'][1] = $anEvent['DTEND'];

                                        // Exclusions
                                        $isExcluded = array_filter($exdates, function ($exdate) use ($anEvent, $monthRecurringOffset) {
                                            return self::isExdateMatch($exdate, $anEvent, $monthRecurringOffset);
                                        });

                                        if (isset($anEvent['UID'])) {
                                            $searchDate = $anEvent['DTSTART'];
                                            if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                                                $searchDate = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent['DTSTART_array'][0]['TZID']) . $searchDate;
                                            }

                                            if (isset($this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                $searchDateUtc = $this->iCalDateToUnixTimestamp($searchDate, true, true);
                                                if (in_array($searchDateUtc, $this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                    $isExcluded = true;
                                                }
                                            }
                                        }

                                        if (!$isExcluded) {
                                            $anEvent            = $this->processEventIcalDateTime($anEvent);
                                            $recurrenceEvents[] = $anEvent;
                                            $this->eventCount++;

                                            // If RRULE[COUNT] is reached then break
                                            if (isset($rrules['COUNT'])) {
                                                $countNb++;

                                                if ($countNb >= $countOrig) {
                                                    break 2;
                                                }
                                            }
                                        }
                                    }

                                    // Move forwards
                                    $recurringTimestamp = strtotime($offset, $recurringTimestamp);
                                }
                            } elseif (isset($rrules['BYDAY']) && $rrules['BYDAY'] !== '') {
                                while ($recurringTimestamp <= $until) {
                                    $monthRecurringTimestamp = $recurringTimestamp;

                                    // Adjust time zone from initial event
                                    $monthRecurringOffset = 0;

                                    if ($this->useTimeZoneWithRRules) {
                                        $recurringTimeZone = \DateTime::createFromFormat(self::UNIX_FORMAT, $monthRecurringTimestamp);
                                        $recurringTimeZone->setTimezone($initialStart->getTimezone());
                                        $monthRecurringOffset = $recurringTimeZone->getOffset();
                                        $monthRecurringTimestamp += $monthRecurringOffset;
                                    }

                                    $eventStartDesc = "{$this->convertDayOrdinalToPositive($dayNumber, $weekday, $monthRecurringTimestamp)} {$this->weekdays[$weekday]} of "
                                        . date(self::DATE_TIME_FORMAT_PRETTY, $monthRecurringTimestamp);
                                    $eventStartTimestamp = strtotime($eventStartDesc);

                                    if (intval($rrules['BYDAY']) === 0) {
                                        $lastDayDesc = "last {$this->weekdays[$weekday]} of "
                                            . date(self::DATE_TIME_FORMAT_PRETTY, $monthRecurringTimestamp);
                                    } else {
                                        $lastDayDesc = "{$this->convertDayOrdinalToPositive($dayNumber, $weekday, $monthRecurringTimestamp)} {$this->weekdays[$weekday]} of "
                                            . date(self::DATE_TIME_FORMAT_PRETTY, $monthRecurringTimestamp);
                                    }

                                    $lastDayTimestamp = strtotime($lastDayDesc);

                                    do {
                                        // Prevent 5th day of a month from showing up on the next month
                                        // If BYDAY and the event falls outside the current month, skip the event

                                        $compareCurrentMonth = date('F', $monthRecurringTimestamp);
                                        $compareEventMonth   = date('F', $eventStartTimestamp);

                                        if ($compareCurrentMonth !== $compareEventMonth) {
                                            $monthRecurringTimestamp = strtotime($offset, $monthRecurringTimestamp);
                                            continue;
                                        }

                                        if ($eventStartTimestamp > $startTimestamp && $eventStartTimestamp <= $until) {
                                            $anEvent['DTSTART'] = date(self::DATE_TIME_FORMAT, $eventStartTimestamp) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');
                                            $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                            $anEvent['DTSTART_array'][2] = $eventStartTimestamp;
                                            $anEvent['DTEND_array']      = $anEvent['DTSTART_array'];
                                            $anEvent['DTEND_array'][2]  += $eventTimestampOffset;
                                            $anEvent['DTEND'] = date(
                                                    self::DATE_TIME_FORMAT,
                                                    $anEvent['DTEND_array'][2]
                                                ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');
                                            $anEvent['DTEND_array'][1] = $anEvent['DTEND'];

                                            // Exclusions
                                            $isExcluded = array_filter($exdates, function ($exdate) use ($anEvent, $monthRecurringOffset) {
                                                return self::isExdateMatch($exdate, $anEvent, $monthRecurringOffset);
                                            });

                                            if (isset($anEvent['UID'])) {
                                                $searchDate = $anEvent['DTSTART'];
                                                if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                                                    $searchDate = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent['DTSTART_array'][0]['TZID']) . $searchDate;
                                                }

                                                if (isset($this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                    $searchDateUtc = $this->iCalDateToUnixTimestamp($searchDate, true, true);
                                                    if (in_array($searchDateUtc, $this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                        $isExcluded = true;
                                                    }
                                                }
                                            }

                                            if (!$isExcluded) {
                                                $anEvent            = $this->processEventIcalDateTime($anEvent);
                                                $recurrenceEvents[] = $anEvent;
                                                $this->eventCount++;

                                                // If RRULE[COUNT] is reached then break
                                                if (isset($rrules['COUNT'])) {
                                                    $countNb++;

                                                    if ($countNb >= $countOrig) {
                                                        break 2;
                                                    }
                                                }
                                            }
                                        }

                                        if (isset($rrules['BYSETPOS'])) {
                                            // BYSETPOS is defined so skip
                                            // looping through each week
                                            $lastDayTimestamp = $eventStartTimestamp;
                                        }

                                        $eventStartTimestamp += self::SECONDS_IN_A_WEEK;
                                    } while ($eventStartTimestamp <= $lastDayTimestamp);

                                    // Move forwards
                                    $recurringTimestamp = strtotime($offset, $recurringTimestamp);
                                }
                            }

                            $recurrenceEvents    = $this->trimToRecurrenceCount($rrules, $recurrenceEvents);
                            $allRecurrenceEvents = array_merge($allRecurrenceEvents, $recurrenceEvents);
                            $recurrenceEvents    = array(); // Reset
                        break;

                        case 'YEARLY':
                            // Create offset
                            $recurringTimestamp = $startTimestamp;
                            $offset = "+{$interval} year";

                            // Deal with BYMONTH
                            if (isset($rrules['BYMONTH']) && $rrules['BYMONTH'] !== '') {
                                $bymonths = explode(',', $rrules['BYMONTH']);
                            } else {
                                $bymonths = array();
                            }

                            // Check if BYDAY rule exists
                            if (isset($rrules['BYDAY']) && $rrules['BYDAY'] !== '') {
                                while ($recurringTimestamp <= $until) {
                                    $yearRecurringTimestamp = $recurringTimestamp;

                                    // Adjust time zone from initial event
                                    $yearRecurringOffset = 0;

                                    if ($this->useTimeZoneWithRRules) {
                                        $recurringTimeZone = \DateTime::createFromFormat(self::UNIX_FORMAT, $yearRecurringTimestamp);
                                        $recurringTimeZone->setTimezone($initialStart->getTimezone());
                                        $yearRecurringOffset = $recurringTimeZone->getOffset();
                                        $yearRecurringTimestamp += $yearRecurringOffset;
                                    }

                                    foreach ($bymonths as $bymonth) {
                                        $eventStartDesc = "{$this->convertDayOrdinalToPositive($dayNumber, $weekday, $yearRecurringTimestamp)} {$this->weekdays[$weekday]}"
                                            . " of {$this->monthNames[$bymonth]} "
                                            . gmdate('Y H:i:s', $yearRecurringTimestamp);
                                        $eventStartTimestamp = strtotime($eventStartDesc);

                                        if (intval($rrules['BYDAY']) === 0) {
                                            $lastDayDesc = "last {$this->weekdays[$weekday]}"
                                                . " of {$this->monthNames[$bymonth]} "
                                                . gmdate('Y H:i:s', $yearRecurringTimestamp);
                                        } else {
                                            $lastDayDesc = "{$this->convertDayOrdinalToPositive($dayNumber, $weekday, $yearRecurringTimestamp)} {$this->weekdays[$weekday]}"
                                                . " of {$this->monthNames[$bymonth]} "
                                                . gmdate('Y H:i:s', $yearRecurringTimestamp);
                                        }

                                        $lastDayTimestamp = strtotime($lastDayDesc);

                                        do {
                                            if ($eventStartTimestamp > $startTimestamp && $eventStartTimestamp <= $until) {
                                                $anEvent['DTSTART'] = date(self::DATE_TIME_FORMAT, $eventStartTimestamp) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');
                                                $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                                $anEvent['DTSTART_array'][2] = $eventStartTimestamp;
                                                $anEvent['DTEND_array']      = $anEvent['DTSTART_array'];
                                                $anEvent['DTEND_array'][2]  += $eventTimestampOffset;
                                                $anEvent['DTEND'] = date(
                                                        self::DATE_TIME_FORMAT,
                                                        $anEvent['DTEND_array'][2]
                                                    ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');
                                                $anEvent['DTEND_array'][1] = $anEvent['DTEND'];

                                                // Exclusions
                                                $isExcluded = array_filter($exdates, function ($exdate) use ($anEvent, $yearRecurringOffset) {
                                                    return self::isExdateMatch($exdate, $anEvent, $yearRecurringOffset);
                                                });

                                                if (isset($anEvent['UID'])) {
                                                    $searchDate = $anEvent['DTSTART'];
                                                    if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                                                        $searchDate = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent['DTSTART_array'][0]['TZID']) . $searchDate;
                                                    }

                                                    if (isset($this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                        $searchDateUtc = $this->iCalDateToUnixTimestamp($searchDate, true, true);
                                                        if (in_array($searchDateUtc, $this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                            $isExcluded = true;
                                                        }
                                                    }
                                                }

                                                if (!$isExcluded) {
                                                    $anEvent            = $this->processEventIcalDateTime($anEvent);
                                                    $recurrenceEvents[] = $anEvent;
                                                    $this->eventCount++;

                                                    // If RRULE[COUNT] is reached then break
                                                    if (isset($rrules['COUNT'])) {
                                                        $countNb++;

                                                        if ($countNb >= $countOrig) {
                                                            break 3;
                                                        }
                                                    }
                                                }
                                            }

                                            $eventStartTimestamp += self::SECONDS_IN_A_WEEK;
                                        } while ($eventStartTimestamp <= $lastDayTimestamp);
                                    }

                                    // Move forwards
                                    $recurringTimestamp = strtotime($offset, $recurringTimestamp);
                                }
                            } else {
                                $day = $initialStart->format('d');

                                // Step through years
                                while ($recurringTimestamp <= $until) {
                                    $yearRecurringTimestamp = $recurringTimestamp;

                                    // Adjust time zone from initial event
                                    $yearRecurringOffset = 0;
                                    if ($this->useTimeZoneWithRRules) {
                                        $recurringTimeZone = \DateTime::createFromFormat(self::UNIX_FORMAT, $yearRecurringTimestamp);
                                        $recurringTimeZone->setTimezone($initialStart->getTimezone());
                                        $yearRecurringOffset = $recurringTimeZone->getOffset();
                                        $yearRecurringTimestamp += $yearRecurringOffset;
                                    }

                                    $eventStartDescs = array();
                                    if (isset($rrules['BYMONTH']) && $rrules['BYMONTH'] !== '') {
                                        foreach ($bymonths as $bymonth) {
                                            array_push($eventStartDescs, "{$day} {$this->monthNames[$bymonth]} " . gmdate('Y H:i:s', $yearRecurringTimestamp));
                                        }
                                    } else {
                                        array_push($eventStartDescs, $day . gmdate(self::DATE_TIME_FORMAT_PRETTY, $yearRecurringTimestamp));
                                    }

                                    foreach ($eventStartDescs as $eventStartDesc) {
                                        $eventStartTimestamp = strtotime($eventStartDesc);

                                        if ($eventStartTimestamp > $startTimestamp && $eventStartTimestamp <= $until) {
                                            $anEvent['DTSTART'] = date(self::DATE_TIME_FORMAT, $eventStartTimestamp) . ($isAllDayEvent || ($initialStartTimeZoneName === 'Z') ? 'Z' : '');
                                            $anEvent['DTSTART_array'][1] = $anEvent['DTSTART'];
                                            $anEvent['DTSTART_array'][2] = $eventStartTimestamp;
                                            $anEvent['DTEND_array']      = $anEvent['DTSTART_array'];
                                            $anEvent['DTEND_array'][2]  += $eventTimestampOffset;
                                            $anEvent['DTEND'] = date(
                                                    self::DATE_TIME_FORMAT,
                                                    $anEvent['DTEND_array'][2]
                                                ) . ($isAllDayEvent || ($initialEndTimeZoneName === 'Z') ? 'Z' : '');
                                            $anEvent['DTEND_array'][1] = $anEvent['DTEND'];

                                            // Exclusions
                                            $isExcluded = array_filter($exdates, function ($exdate) use ($anEvent, $yearRecurringOffset) {
                                                return self::isExdateMatch($exdate, $anEvent, $yearRecurringOffset);
                                            });

                                            if (isset($anEvent['UID'])) {
                                                $searchDate = $anEvent['DTSTART'];
                                                if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                                                    $searchDate = sprintf(self::ICAL_DATE_TIME_TEMPLATE, $anEvent['DTSTART_array'][0]['TZID']) . $searchDate;
                                                }

                                                if (isset($this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                    $searchDateUtc = $this->iCalDateToUnixTimestamp($searchDate, true, true);
                                                    if (in_array($searchDateUtc, $this->alteredRecurrenceInstances[$anEvent['UID']])) {
                                                        $isExcluded = true;
                                                    }
                                                }
                                            }

                                            if (!$isExcluded) {
                                                $anEvent            = $this->processEventIcalDateTime($anEvent);
                                                $recurrenceEvents[] = $anEvent;
                                                $this->eventCount++;

                                                // If RRULE[COUNT] is reached then break
                                                if (isset($rrules['COUNT'])) {
                                                    $countNb++;

                                                    if ($countNb >= $countOrig) {
                                                        break 2;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // Move forwards
                                    $recurringTimestamp = strtotime($offset, $recurringTimestamp);
                                }
                            }

                            $recurrenceEvents    = $this->trimToRecurrenceCount($rrules, $recurrenceEvents);
                            $allRecurrenceEvents = array_merge($allRecurrenceEvents, $recurrenceEvents);
                            $recurrenceEvents    = array(); // Reset
                        break;
                    }
                }
            }

            $events = array_merge($events, $allRecurrenceEvents);

            $this->cal['VEVENT'] = $events;
        }
    }

    /**
     * Processes date conversions using the time zone
     *
     * Add keys `DTSTART_tz` and `DTEND_tz` to each Event
     * These keys contain dates adapted to the calendar
     * time zone depending on the event `TZID`.
     *
     * @return void
     */
    protected function processDateConversions()
    {
        $events = (isset($this->cal['VEVENT'])) ? $this->cal['VEVENT'] : array();

        if (!empty($events)) {
            foreach ($events as $key => $anEvent) {
                if (!$this->isValidDate($anEvent['DTSTART'])) {
                    unset($events[$key]);
                    $this->eventCount--;

                    continue;
                }

                if ($this->useTimeZoneWithRRules && isset($anEvent['RRULE_array'][2]) && $anEvent['RRULE_array'][2] === self::RECURRENCE_EVENT) {
                    $events[$key]['DTSTART_tz'] = $anEvent['DTSTART'];
                    $events[$key]['DTEND_tz']   = isset($anEvent['DTEND']) ? $anEvent['DTEND'] : $anEvent['DTSTART'];
                } else {
                    $events[$key]['DTSTART_tz'] = $this->iCalDateWithTimeZone($anEvent, 'DTSTART');

                    if ($this->iCalDateWithTimeZone($anEvent, 'DTEND')) {
                        $events[$key]['DTEND_tz'] = $this->iCalDateWithTimeZone($anEvent, 'DTEND');
                    } elseif ($this->iCalDateWithTimeZone($anEvent, 'DURATION')) {
                        $events[$key]['DTEND_tz'] = $this->iCalDateWithTimeZone($anEvent, 'DURATION');
                    } elseif ($this->iCalDateWithTimeZone($anEvent, 'DTSTART')) {
                        $events[$key]['DTEND_tz'] = $this->iCalDateWithTimeZone($anEvent, 'DTSTART');
                    }
                }
            }

            $this->cal['VEVENT'] = $events;
        }
    }

    /**
     * Extends the `{DTSTART|DTEND|RECURRENCE-ID}_array`
     * array to include an iCal date time for each event
     * (`TZID=Timezone:YYYYMMDD[T]HHMMSS`)
     *
     * @param  array   $event
     * @param  integer $index
     * @return array
     */
    protected function processEventIcalDateTime(array $event, $index = 3)
    {
        $calendarTimeZone = $this->calendarTimeZone(true);

        foreach (array('DTSTART', 'DTEND', 'RECURRENCE-ID') as $type) {
            if (isset($event["{$type}_array"])) {
                $timeZone = (isset($event["{$type}_array"][0]['TZID'])) ? $event["{$type}_array"][0]['TZID'] : $calendarTimeZone;
                $event["{$type}_array"][$index] = ((is_null($timeZone)) ? '' : sprintf(self::ICAL_DATE_TIME_TEMPLATE, $timeZone)) . $event["{$type}_array"][1];
            }
        }

        return $event;
    }

    /**
     * Returns an array of Events.
     * Every event is a class with the event
     * details being properties within it.
     *
     * @return array
     */
    public function events()
    {
        $array = $this->cal;
        $array = isset($array['VEVENT']) ? $array['VEVENT'] : array();
        $events = array();

        if (!empty($array)) {
            foreach ($array as $event) {
                $events[] = new Event($event);
            }
        }

        return $events;
    }

    /**
     * Returns the calendar name
     *
     * @return string
     */
    public function calendarName()
    {
        return isset($this->cal['VCALENDAR']['X-WR-CALNAME']) ? $this->cal['VCALENDAR']['X-WR-CALNAME'] : '';
    }

    /**
     * Returns the calendar description
     *
     * @return string
     */
    public function calendarDescription()
    {
        return isset($this->cal['VCALENDAR']['X-WR-CALDESC']) ? $this->cal['VCALENDAR']['X-WR-CALDESC'] : '';
    }

    /**
     * Returns the calendar time zone
     *
     * @param  boolean $ignoreUtc
     * @return string
     */
    public function calendarTimeZone($ignoreUtc = false)
    {
        if (isset($this->cal['VCALENDAR']['X-WR-TIMEZONE'])) {
            $timeZone = $this->cal['VCALENDAR']['X-WR-TIMEZONE'];
        } elseif (isset($this->cal['VTIMEZONE']['TZID'])) {
            $timeZone = $this->cal['VTIMEZONE']['TZID'];
        } else {
            $timeZone = $this->defaultTimeZone;
        }

        // Use default time zone if the calendar's is invalid
        if ($this->isValidIanaTimeZoneId($timeZone) === false) {
            // phpcs:ignore CustomPHPCS.ControlStructures.AssignmentInCondition.Warning
            if (($timeZone = $this->isValidCldrTimeZoneId($timeZone, true)) === false) {
                $timeZone = $this->defaultTimeZone;
            }
        }

        if ($ignoreUtc && strtoupper($timeZone) === self::TIME_ZONE_UTC) {
            return null;
        }

        return $timeZone;
    }

    /**
     * Returns an array of arrays with all free/busy events.
     * Every event is an associative array and each property
     * is an element it.
     *
     * @return array
     */
    public function freeBusyEvents()
    {
        $array = $this->cal;

        return isset($array['VFREEBUSY']) ? $array['VFREEBUSY'] : [];
    }

    /**
     * Returns a boolean value whether the
     * current calendar has events or not
     *
     * @return boolean
     */
    public function hasEvents()
    {
        return (count($this->events()) > 0) ?: false;
    }

    /**
     * Returns a sorted array of the events in a given range,
     * or an empty array if no events exist in the range.
     *
     * Events will be returned if the start or end date is contained within the
     * range (inclusive), or if the event starts before and end after the range.
     *
     * If a start date is not specified or of a valid format, then the start
     * of the range will default to the current time and date of the server.
     *
     * If an end date is not specified or of a valid format, then the end of
     * the range will default to the current time and date of the server,
     * plus 20 years.
     *
     * Note that this function makes use of Unix timestamps. This might be a
     * problem for events on, during, or after 29 Jan 2038.
     * See https://en.wikipedia.org/wiki/Unix_time#Representing_the_number
     *
     * @param  string|null $rangeStart
     * @param  string|null $rangeEnd
     * @return array
     * @throws \Exception
     */
    public function eventsFromRange($rangeStart = null, $rangeEnd = null)
    {
        // Sort events before processing range
        $events = $this->sortEventsWithOrder($this->events(), SORT_ASC);

        if (empty($events)) {
            return array();
        }

        $extendedEvents = array();

        if (!is_null($rangeStart)) {
            try {
                $rangeStart = new \DateTime($rangeStart, new \DateTimeZone($this->defaultTimeZone));
            } catch (\Exception $e) {
                error_log("ICal::eventsFromRange: Invalid date passed ({$rangeStart})");
                $rangeStart = false;
            }
        } else {
            $rangeStart = new \DateTime('now', new \DateTimeZone($this->defaultTimeZone));
        }

        if (!is_null($rangeEnd)) {
            try {
                $rangeEnd = new \DateTime($rangeEnd, new \DateTimeZone($this->defaultTimeZone));
            } catch (\Exception $e) {
                error_log("ICal::eventsFromRange: Invalid date passed ({$rangeEnd})");
                $rangeEnd = false;
            }
        } else {
            $rangeEnd = new \DateTime('now', new \DateTimeZone($this->defaultTimeZone));
            $rangeEnd->modify('+20 years');
        }

        // If start and end are identical and are dates with no times...
        if ($rangeEnd->format('His') == 0 && $rangeStart->getTimestamp() == $rangeEnd->getTimestamp()) {
            $rangeEnd->modify('+1 day');
        }

        $rangeStart = $rangeStart->getTimestamp();
        $rangeEnd   = $rangeEnd->getTimestamp();

        foreach ($events as $anEvent) {
            $eventStart = $anEvent->dtstart_array[2];
            $eventEnd   = (isset($anEvent->dtend_array[2])) ? $anEvent->dtend_array[2] : null;

            if (($eventStart >= $rangeStart && $eventStart < $rangeEnd)         // Event start date contained in the range
                || ($eventEnd !== null
                    && (
                        ($eventEnd > $rangeStart && $eventEnd <= $rangeEnd)     // Event end date contained in the range
                        || ($eventStart < $rangeStart && $eventEnd > $rangeEnd) // Event starts before and finishes after range
                    )
                )
            ) {
                $extendedEvents[] = $anEvent;
            }
        }

        if (empty($extendedEvents)) {
            return array();
        }

        return $extendedEvents;
    }

    /**
     * Returns a sorted array of the events following a given string,
     * or `false` if no events exist in the range.
     *
     * @param  string $interval
     * @return array
     */
    public function eventsFromInterval($interval)
    {
        $rangeStart = new \DateTime('now', new \DateTimeZone($this->defaultTimeZone));
        $rangeEnd   = new \DateTime('now', new \DateTimeZone($this->defaultTimeZone));

        $dateInterval = \DateInterval::createFromDateString($interval);
        $rangeEnd->add($dateInterval);

        return $this->eventsFromRange($rangeStart->format('Y-m-d'), $rangeEnd->format('Y-m-d'));
    }

    /**
     * Sorts events based on a given sort order
     *
     * @param  array   $events
     * @param  integer $sortOrder Either SORT_ASC, SORT_DESC, SORT_REGULAR, SORT_NUMERIC, SORT_STRING
     * @return array
     */
    public function sortEventsWithOrder(array $events, $sortOrder = SORT_ASC)
    {
        $extendedEvents = array();
        $timestamp      = array();

        foreach ($events as $key => $anEvent) {
            $extendedEvents[] = $anEvent;
            $timestamp[$key]  = $anEvent->dtstart_array[2];
        }

        array_multisort($timestamp, $sortOrder, $extendedEvents);

        return $extendedEvents;
    }

    /**
     * Checks if a time zone is valid (IANA or CLDR)
     *
     * @param  string $timeZone
     * @return boolean
     */
    protected function isValidTimeZoneId($timeZone)
    {
        return ($this->isValidIanaTimeZoneId($timeZone) !== false || $this->isValidCldrTimeZoneId($timeZone) !== false);
    }

    /**
     * Checks if a time zone is a valid IANA time zone
     *
     * @param  string $timeZone
     * @return boolean
     */
    protected function isValidIanaTimeZoneId($timeZone)
    {
        if (in_array($timeZone, $this->validTimeZones)) {
            return true;
        }

        $valid = array();
        $tza   = timezone_abbreviations_list();

        foreach ($tza as $zone) {
            foreach ($zone as $item) {
                $valid[$item['timezone_id']] = true;
            }
        }

        unset($valid['']);

        if (isset($valid[$timeZone]) || in_array($timeZone, timezone_identifiers_list(\DateTimeZone::ALL_WITH_BC))) {
            $this->validTimeZones[] = $timeZone;

            return true;
        }

        return false;
    }

    /**
     * Checks if a time zone is a valid CLDR time zone
     *
     * @param  string  $timeZone
     * @param  boolean $doConversion
     * @return boolean|string
     */
    public function isValidCldrTimeZoneId($timeZone, $doConversion = false)
    {
        $timeZone = html_entity_decode($timeZone);

        $cldrTimeZones = array(
            '(UTC-12:00) International Date Line West'                      => 'Etc/GMT+12',
            '(UTC-11:00) Coordinated Universal Time-11'                     => 'Etc/GMT+11',
            '(UTC-10:00) Hawaii'                                            => 'Pacific/Honolulu',
            '(UTC-09:00) Alaska'                                            => 'America/Anchorage',
            '(UTC-08:00) Pacific Time (US & Canada)'                        => 'America/Los_Angeles',
            '(UTC-07:00) Arizona'                                           => 'America/Phoenix',
            '(UTC-07:00) Chihuahua, La Paz, Mazatlan'                       => 'America/Chihuahua',
            '(UTC-07:00) Mountain Time (US & Canada)'                       => 'America/Denver',
            '(UTC-06:00) Central America'                                   => 'America/Guatemala',
            '(UTC-06:00) Central Time (US & Canada)'                        => 'America/Chicago',
            '(UTC-06:00) Guadalajara, Mexico City, Monterrey'               => 'America/Mexico_City',
            '(UTC-06:00) Saskatchewan'                                      => 'America/Regina',
            '(UTC-05:00) Bogota, Lima, Quito, Rio Branco'                   => 'America/Bogota',
            '(UTC-05:00) Chetumal'                                          => 'America/Cancun',
            '(UTC-05:00) Eastern Time (US & Canada)'                        => 'America/New_York',
            '(UTC-05:00) Indiana (East)'                                    => 'America/Indianapolis',
            '(UTC-04:00) Asuncion'                                          => 'America/Asuncion',
            '(UTC-04:00) Atlantic Time (Canada)'                            => 'America/Halifax',
            '(UTC-04:00) Caracas'                                           => 'America/Caracas',
            '(UTC-04:00) Cuiaba'                                            => 'America/Cuiaba',
            '(UTC-04:00) Georgetown, La Paz, Manaus, San Juan'              => 'America/La_Paz',
            '(UTC-04:00) Santiago'                                          => 'America/Santiago',
            '(UTC-03:30) Newfoundland'                                      => 'America/St_Johns',
            '(UTC-03:00) Brasilia'                                          => 'America/Sao_Paulo',
            '(UTC-03:00) Cayenne, Fortaleza'                                => 'America/Cayenne',
            '(UTC-03:00) City of Buenos Aires'                              => 'America/Buenos_Aires',
            '(UTC-03:00) Greenland'                                         => 'America/Godthab',
            '(UTC-03:00) Montevideo'                                        => 'America/Montevideo',
            '(UTC-03:00) Salvador'                                          => 'America/Bahia',
            '(UTC-02:00) Coordinated Universal Time-02'                     => 'Etc/GMT+2',
            '(UTC-01:00) Azores'                                            => 'Atlantic/Azores',
            '(UTC-01:00) Cabo Verde Is.'                                    => 'Atlantic/Cape_Verde',
            '(UTC) Coordinated Universal Time'                              => 'Etc/GMT',
            '(UTC+00:00) Casablanca'                                        => 'Africa/Casablanca',
            '(UTC+00:00) Dublin, Edinburgh, Lisbon, London'                 => 'Europe/London',
            '(UTC+00:00) Monrovia, Reykjavik'                               => 'Atlantic/Reykjavik',
            '(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna'  => 'Europe/Berlin',
            '(UTC+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague' => 'Europe/Budapest',
            '(UTC+01:00) Brussels, Copenhagen, Madrid, Paris'               => 'Europe/Paris',
            '(UTC+01:00) Sarajevo, Skopje, Warsaw, Zagreb'                  => 'Europe/Warsaw',
            '(UTC+01:00) West Central Africa'                               => 'Africa/Lagos',
            '(UTC+02:00) Amman'                                             => 'Asia/Amman',
            '(UTC+02:00) Athens, Bucharest'                                 => 'Europe/Bucharest',
            '(UTC+02:00) Beirut'                                            => 'Asia/Beirut',
            '(UTC+02:00) Cairo'                                             => 'Africa/Cairo',
            '(UTC+02:00) Chisinau'                                          => 'Europe/Chisinau',
            '(UTC+02:00) Damascus'                                          => 'Asia/Damascus',
            '(UTC+02:00) Harare, Pretoria'                                  => 'Africa/Johannesburg',
            '(UTC+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius'     => 'Europe/Kiev',
            '(UTC+02:00) Jerusalem'                                         => 'Asia/Jerusalem',
            '(UTC+02:00) Kaliningrad'                                       => 'Europe/Kaliningrad',
            '(UTC+02:00) Tripoli'                                           => 'Africa/Tripoli',
            '(UTC+02:00) Windhoek'                                          => 'Africa/Windhoek',
            '(UTC+03:00) Baghdad'                                           => 'Asia/Baghdad',
            '(UTC+03:00) Istanbul'                                          => 'Europe/Istanbul',
            '(UTC+03:00) Kuwait, Riyadh'                                    => 'Asia/Riyadh',
            '(UTC+03:00) Minsk'                                             => 'Europe/Minsk',
            '(UTC+03:00) Moscow, St. Petersburg, Volgograd'                 => 'Europe/Moscow',
            '(UTC+03:00) Nairobi'                                           => 'Africa/Nairobi',
            '(UTC+03:30) Tehran'                                            => 'Asia/Tehran',
            '(UTC+04:00) Abu Dhabi, Muscat'                                 => 'Asia/Dubai',
            '(UTC+04:00) Baku'                                              => 'Asia/Baku',
            '(UTC+04:00) Izhevsk, Samara'                                   => 'Europe/Samara',
            '(UTC+04:00) Port Louis'                                        => 'Indian/Mauritius',
            '(UTC+04:00) Tbilisi'                                           => 'Asia/Tbilisi',
            '(UTC+04:00) Yerevan'                                           => 'Asia/Yerevan',
            '(UTC+04:30) Kabul'                                             => 'Asia/Kabul',
            '(UTC+05:00) Ashgabat, Tashkent'                                => 'Asia/Tashkent',
            '(UTC+05:00) Ekaterinburg'                                      => 'Asia/Yekaterinburg',
            '(UTC+05:00) Islamabad, Karachi'                                => 'Asia/Karachi',
            '(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi'               => 'Asia/Calcutta',
            '(UTC+05:30) Sri Jayawardenepura'                               => 'Asia/Colombo',
            '(UTC+05:45) Kathmandu'                                         => 'Asia/Katmandu',
            '(UTC+06:00) Astana'                                            => 'Asia/Almaty',
            '(UTC+06:00) Dhaka'                                             => 'Asia/Dhaka',
            '(UTC+06:30) Yangon (Rangoon)'                                  => 'Asia/Rangoon',
            '(UTC+07:00) Bangkok, Hanoi, Jakarta'                           => 'Asia/Bangkok',
            '(UTC+07:00) Krasnoyarsk'                                       => 'Asia/Krasnoyarsk',
            '(UTC+07:00) Novosibirsk'                                       => 'Asia/Novosibirsk',
            '(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi'             => 'Asia/Shanghai',
            '(UTC+08:00) Irkutsk'                                           => 'Asia/Irkutsk',
            '(UTC+08:00) Kuala Lumpur, Singapore'                           => 'Asia/Singapore',
            '(UTC+08:00) Perth'                                             => 'Australia/Perth',
            '(UTC+08:00) Taipei'                                            => 'Asia/Taipei',
            '(UTC+08:00) Ulaanbaatar'                                       => 'Asia/Ulaanbaatar',
            '(UTC+09:00) Osaka, Sapporo, Tokyo'                             => 'Asia/Tokyo',
            '(UTC+09:00) Pyongyang'                                         => 'Asia/Pyongyang',
            '(UTC+09:00) Seoul'                                             => 'Asia/Seoul',
            '(UTC+09:00) Yakutsk'                                           => 'Asia/Yakutsk',
            '(UTC+09:30) Adelaide'                                          => 'Australia/Adelaide',
            '(UTC+09:30) Darwin'                                            => 'Australia/Darwin',
            '(UTC+10:00) Brisbane'                                          => 'Australia/Brisbane',
            '(UTC+10:00) Canberra, Melbourne, Sydney'                       => 'Australia/Sydney',
            '(UTC+10:00) Guam, Port Moresby'                                => 'Pacific/Port_Moresby',
            '(UTC+10:00) Hobart'                                            => 'Australia/Hobart',
            '(UTC+10:00) Vladivostok'                                       => 'Asia/Vladivostok',
            '(UTC+11:00) Chokurdakh'                                        => 'Asia/Srednekolymsk',
            '(UTC+11:00) Magadan'                                           => 'Asia/Magadan',
            '(UTC+11:00) Solomon Is., New Caledonia'                        => 'Pacific/Guadalcanal',
            '(UTC+12:00) Anadyr, Petropavlovsk-Kamchatsky'                  => 'Asia/Kamchatka',
            '(UTC+12:00) Auckland, Wellington'                              => 'Pacific/Auckland',
            '(UTC+12:00) Coordinated Universal Time+12'                     => 'Etc/GMT-12',
            '(UTC+12:00) Fiji'                                              => 'Pacific/Fiji',
            "(UTC+13:00) Nuku'alofa"                                        => 'Pacific/Tongatapu',
            '(UTC+13:00) Samoa'                                             => 'Pacific/Apia',
            '(UTC+14:00) Kiritimati Island'                                 => 'Pacific/Kiritimati',
        );

        if (array_key_exists($timeZone, $cldrTimeZones)) {
            if ($doConversion) {
                return $cldrTimeZones[$timeZone];
            } else {
                return true;
            }
        }

        return false;
    }

    /**
     * Parses a duration and applies it to a date
     *
     * @param  string $date
     * @param  string $duration
     * @param  string $format
     * @return integer|\DateTime
     */
    protected function parseDuration($date, $duration, $format = self::UNIX_FORMAT)
    {
        $dateTime = date_create($date);
        $dateTime->modify($duration->y . ' year');
        $dateTime->modify($duration->m . ' month');
        $dateTime->modify($duration->d . ' day');
        $dateTime->modify($duration->h . ' hour');
        $dateTime->modify($duration->i . ' minute');
        $dateTime->modify($duration->s . ' second');

        if (is_null($format)) {
            $output = $dateTime;
        } else {
            if ($format === self::UNIX_FORMAT) {
                $output = $dateTime->getTimestamp();
            } else {
                $output = $dateTime->format($format);
            }
        }

        return $output;
    }

    /**
     * Gets the number of days between a start and end date
     *
     * @param  integer $days
     * @param  integer $start
     * @param  integer $end
     * @return integer
     */
    protected function numberOfDays($days, $start, $end)
    {
        $w    = array(date('w', $start), date('w', $end));
        $base = floor(($end - $start) / self::SECONDS_IN_A_WEEK);
        $sum  = 0;

        for ($day = 0; $day < 7; ++$day) {
            if ($days & pow(2, $day)) {
                $sum += $base + (($w[0] > $w[1]) ? $w[0] <= $day || $day <= $w[1] : $w[0] <= $day && $day <= $w[1]);
            }
        }

        return $sum;
    }

    /**
     * Converts a negative day ordinal to
     * its equivalent positive form
     *
     * @param  integer $dayNumber
     * @param  integer $weekday
     * @param  integer|\DateTime $timestamp
     * @return string
     */
    protected function convertDayOrdinalToPositive($dayNumber, $weekday, $timestamp)
    {
        $dayNumber = empty($dayNumber) ? 1 : $dayNumber; // Returns 0 when no number defined in BYDAY

        $dayOrdinals = $this->dayOrdinals;

        // We only care about negative BYDAY values
        if ($dayNumber >= 1) {
            return $dayOrdinals[$dayNumber];
        }

        $timestamp = (is_object($timestamp)) ? $timestamp : \DateTime::createFromFormat(self::UNIX_FORMAT, $timestamp);
        $start     = strtotime('first day of ' . $timestamp->format(self::DATE_TIME_FORMAT_PRETTY));
        $end       = strtotime('last day of ' . $timestamp->format(self::DATE_TIME_FORMAT_PRETTY));

        // Used with pow(2, X) so pow(2, 4) is THURSDAY
        $weekdays = array_flip(array_keys($this->weekdays));

        $numberOfDays = $this->numberOfDays(pow(2, $weekdays[$weekday]), $start, $end);

        // Create subset
        $dayOrdinals = array_slice($dayOrdinals, 0, $numberOfDays, true);

        // Reverse only the values
        $dayOrdinals = array_combine(array_keys($dayOrdinals), array_reverse(array_values($dayOrdinals)));

        return $dayOrdinals[$dayNumber * -1];
    }

    /**
     * Removes unprintable ASCII and UTF-8 characters
     *
     * @param  string $data
     * @return string
     */
    protected function removeUnprintableChars($data)
    {
        return preg_replace('/[\x00-\x1F\x7F\xA0]/u', '', $data);
    }

    /**
     * Provides a polyfill for PHP 7.2's `mb_chr()`, which is a multibyte safe version of `chr()`.
     * Multibyte safe.
     *
     * @param  integer $code
     * @return string
     */
    protected function mb_chr($code)
    {
        if (function_exists('mb_chr')) {
            return mb_chr($code);
        } else {
            if (0x80 > $code %= 0x200000) {
                $s = chr($code);
            } elseif (0x800 > $code) {
                $s = chr(0xc0 | $code >> 6) . chr(0x80 | $code & 0x3f);
            } elseif (0x10000 > $code) {
                $s = chr(0xe0 | $code >> 12) . chr(0x80 | $code >> 6 & 0x3f) . chr(0x80 | $code & 0x3f);
            } else {
                $s = chr(0xf0 | $code >> 18) . chr(0x80 | $code >> 12 & 0x3f) . chr(0x80 | $code >> 6 & 0x3f) . chr(0x80 | $code & 0x3f);
            }

            return $s;
        }
    }

    /**
     * Replace all occurrences of the search string with the replacement string.
     * Multibyte safe.
     *
     * @param  string|array $search
     * @param  string|array $replace
     * @param  string|array $subject
     * @param  string       $encoding
     * @param  integer      $count
     * @return array|string
     */
    protected static function mb_str_replace($search, $replace, $subject, $encoding = null, &$count = 0)
    {
        if (is_array($subject)) {
            // Call `mb_str_replace()` for each subject in the array, recursively
            foreach ($subject as $key => $value) {
                $subject[$key] = self::mb_str_replace($search, $replace, $value, $encoding, $count);
            }
        } else {
            // Normalize $search and $replace so they are both arrays of the same length
            $searches     = is_array($search) ? array_values($search) : [$search];
            $replacements = is_array($replace) ? array_values($replace) : [$replace];
            $replacements = array_pad($replacements, count($searches), '');

            foreach ($searches as $key => $search) {
                if (is_null($encoding)) {
                    $encoding = mb_detect_encoding($search, 'UTF-8', true);
                }

                $replace   = $replacements[$key];
                $searchLen = mb_strlen($search, $encoding);

                $sb = [];
                while (($offset = mb_strpos($subject, $search, 0, $encoding)) !== false) {
                    $sb[]    = mb_substr($subject, 0, $offset, $encoding);
                    $subject = mb_substr($subject, $offset + $searchLen, null, $encoding);
                    ++$count;
                }

                $sb[]    = $subject;
                $subject = implode($replace, $sb);
            }
        }

        return $subject;
    }

    /**
     * Replaces curly quotes and other special characters
     * with their standard equivalents
     *
     * @param  string $data
     * @return string
     */
    protected function cleanData($data)
    {
        $replacementChars = array(
            "\xe2\x80\x98" => "'",   // ‘
            "\xe2\x80\x99" => "'",   // ’
            "\xe2\x80\x9a" => "'",   // ‚
            "\xe2\x80\x9b" => "'",   // ‛
            "\xe2\x80\x9c" => '"',   // “
            "\xe2\x80\x9d" => '"',   // ”
            "\xe2\x80\x9e" => '"',   // „
            "\xe2\x80\x9f" => '"',   // ‟
            "\xe2\x80\x93" => '-',   // –
            "\xe2\x80\x94" => '--',  // —
            "\xe2\x80\xa6" => '...', // …
            "\xc2\xa0"     => ' ',
        );
        // Replace UTF-8 characters
        $cleanedData = strtr($data, $replacementChars);

        // Replace Windows-1252 equivalents
        $charsToReplace = array_map(function ($code) {
            return $this->mb_chr($code);
        }, array(133, 145, 146, 147, 148, 150, 151, 194));
        $cleanedData = $this->mb_str_replace($charsToReplace, $replacementChars, $cleanedData);

        return $cleanedData;
    }

    /**
     * Parses a list of excluded dates
     * to be applied to an Event
     *
     * @param  array $event
     * @return array
     */
    public function parseExdates(array $event)
    {
        if (empty($event['EXDATE_array'])) {
            return array();
        } else {
            $exdates = $event['EXDATE_array'];
        }

        $output          = array();
        $currentTimeZone = $this->defaultTimeZone;

        foreach ($exdates as $subArray) {
            end($subArray);
            $finalKey = key($subArray);

            foreach ($subArray as $key => $value) {
                if ($key === 'TZID') {
                    $checkTimeZone = $subArray[$key];

                    if ($this->isValidIanaTimeZoneId($checkTimeZone)) {
                        $currentTimeZone = $checkTimeZone;
                    } elseif ($this->isValidCldrTimeZoneId($checkTimeZone)) {
                        $currentTimeZone = $this->isValidCldrTimeZoneId($checkTimeZone, true);
                    } else {
                        $currentTimeZone = $this->defaultTimeZone;
                    }
                } elseif (is_numeric($key)) {
                    $icalDate = $subArray[$key];

                    if (substr($icalDate, -1) === 'Z') {
                        $currentTimeZone = self::TIME_ZONE_UTC;
                    }

                    $output[] = new Carbon($icalDate, $currentTimeZone);

                    if ($key === $finalKey) {
                        // Reset to default
                        $currentTimeZone = $this->defaultTimeZone;
                    }
                }
            }
        }

        return $output;
    }

    /**
     * Checks if a date string is a valid date
     *
     * @param  string $value
     * @return boolean
     * @throws \Exception
     */
    public function isValidDate($value)
    {
        if (!$value) {
            return false;
        }

        try {
            new \DateTime($value);

            return true;
        } catch (\Exception $e) {
            return false;
        }
    }

    /**
     * Checks if a filename exists as a file or URL
     *
     * @param  string $filename
     * @return boolean
     */
    protected function isFileOrUrl($filename)
    {
        return (file_exists($filename) || filter_var($filename, FILTER_VALIDATE_URL)) ?: false;
    }

    /**
     * Reads an entire file or URL into an array
     *
     * @param  string $filename
     * @return array
     * @throws \Exception
     */
    protected function fileOrUrl($filename)
    {
        $options = array();
        if (!empty($this->httpBasicAuth)) {
            $options['http'] = array();

            $username  = $this->httpBasicAuth['username'];
            $password  = $this->httpBasicAuth['password'];
            $basicAuth = base64_encode("{$username}:{$password}");

            $options['http']['header'] = "Authorization: Basic {$basicAuth}";
        }

        $context = stream_context_create($options);

        if (($lines = file($filename, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES, $context)) === false) {
            throw new \Exception("The file path or URL '{$filename}' does not exist.");
        }

        return $lines;
    }

    /**
     * Ensures the recurrence count is enforced against generated recurrence events.
     *
     * @param  array $rrules
     * @param  array $recurrenceEvents
     * @return array
     */
    protected function trimToRecurrenceCount(array $rrules, array $recurrenceEvents)
    {
        if (isset($rrules['COUNT'])) {
            $recurrenceCount = (intval($rrules['COUNT']) - 1);
            $surplusCount    = (sizeof($recurrenceEvents) - $recurrenceCount);

            if ($surplusCount > 0) {
                $recurrenceEvents  = array_slice($recurrenceEvents, 0, $recurrenceCount);
                $this->eventCount -= $surplusCount;
            }
        }

        return $recurrenceEvents;
    }

    /**
     * Checks if an excluded date matches a given date by reconciling time zones.
     *
     * @param  Carbon $exdate
     * @param  array   $anEvent
     * @param  integer $recurringOffset
     * @return boolean
     */
    protected function isExdateMatch($exdate, array $anEvent, $recurringOffset)
    {
        $searchDate = $anEvent['DTSTART'];

        if (substr($searchDate, -1) === 'Z') {
            $timeZone = self::TIME_ZONE_UTC;
        } else {
            if (isset($anEvent['DTSTART_array'][0]['TZID'])) {
                $checkTimeZone = $anEvent['DTSTART_array'][0]['TZID'];

                if ($this->isValidIanaTimeZoneId($checkTimeZone)) {
                    $timeZone = $checkTimeZone;
                } elseif ($this->isValidCldrTimeZoneId($checkTimeZone)) {
                    $timeZone = $this->isValidCldrTimeZoneId($checkTimeZone, true);
                } else {
                    $timeZone = $this->defaultTimeZone;
                }
            } else {
                $timeZone = $this->defaultTimeZone;
            }
        }

        $a = new Carbon($searchDate, $timeZone);
        $b = $exdate->addSeconds($recurringOffset);

        return $a->eq($b);
    }

    /**
     * Replaces non-CLDR Windows time zone ID like 'W. Europe Standard Time' with its IANA equivalent.
     *
     * @param  string $lineWithTzid
     * @return string
     */
    protected function replaceWindowsTimeZoneId($lineWithTzid)
    {
        return str_replace($this->windowsTimeZones, $this->windowsTimeZonesIana, $lineWithTzid);
    }
}
