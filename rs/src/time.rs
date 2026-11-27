//! Dates as the TypeScript's `Date` produces them, for the three places the
//! search uses one: "last synced" (`toLocaleString()`), `--since`, and the
//! day shown beside a page.
//!
//! `Date` reads many notations and has an answer for every time zone setting.
//! Only the plain ones are taken here; `HandBack` is returned for the rest,
//! and the caller hands the invocation to the JavaScript.

use jiff::civil::{Date, DateTime, Time};
use jiff::tz::TimeZone;
use jiff::Timestamp;

use crate::HandBack;

/// A zone by name. The database compiled into the binary is asked, as the
/// JavaScript runtime also carries its own (ICU's).
fn zone(name: &str) -> Option<TimeZone> {
    let zone = jiff::tz::TimeZoneDatabase::bundled().get(name).ok()?;
    // The runtime also takes `asia/tokyo`; only the name as it is written in
    // the database is taken here.
    (zone.iana_name() == Some(name)).then_some(zone)
}

/// The zone `Date` uses for local time: the one `TZ` names, else the
/// system's. `TZ` can also hold a rule (`JST-9`), a path or nothing, which
/// the JavaScript runtime reads its own way.
pub fn local_zone() -> Result<TimeZone, HandBack> {
    match std::env::var_os("TZ") {
        Some(tz) => tz.to_str().and_then(zone).ok_or(HandBack),
        None => TimeZone::try_system().ok().and_then(|system| system.iana_name().and_then(zone)).ok_or(HandBack),
    }
}

/// `new Date(seconds * 1000).toLocaleString()`. The runtime's locale is
/// en-US whatever the user's is (bun does not read LANG or LC_*).
pub fn locale_string(seconds: i64, zone: &TimeZone) -> Result<String, HandBack> {
    let local = zone.to_datetime(Timestamp::from_second(seconds).map_err(|_| HandBack)?);
    if !(1000..=9999).contains(&local.year()) {
        return Err(HandBack);
    }
    let (hour, half) = match local.hour() {
        0 => (12, "AM"),
        hour @ 1..=11 => (hour, "AM"),
        12 => (12, "PM"),
        hour => (hour - 12, "PM"),
    };
    Ok(format!(
        "{}/{}/{}, {}:{:02}:{:02} {}",
        local.month(),
        local.day(),
        local.year(),
        hour,
        local.minute(),
        local.second(),
        half
    ))
}

fn digits(text: &str, count: usize) -> Option<i64> {
    (text.len() == count && text.bytes().all(|b| b.is_ascii_digit())).then(|| text.parse().ok()).flatten()
}

/// `YYYY-MM-DD` naming a day that exists, in the years written with four
/// digits and no sign.
fn calendar_day(text: &str) -> Option<Date> {
    let bytes = text.as_bytes();
    if !text.is_ascii() || bytes.len() != 10 || bytes[4] != b'-' || bytes[7] != b'-' {
        return None;
    }
    let (year, month, day) = (digits(&text[0..4], 4)?, digits(&text[5..7], 2)?, digits(&text[8..10], 2)?);
    if !(1000..=9999).contains(&year) {
        return None;
    }
    // `Date` would turn February 30th into March 2nd; not taken on.
    Date::new(year as i16, month as i8, day as i8).ok()
}

/// `d.toISOString()`
fn iso_string(instant: Timestamp) -> Result<String, HandBack> {
    let utc = TimeZone::UTC.to_datetime(instant);
    if !(1000..=9999).contains(&utc.year()) {
        return Err(HandBack);
    }
    Ok(format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}Z",
        utc.year(),
        utc.month(),
        utc.day(),
        utc.hour(),
        utc.minute(),
        utc.second(),
        utc.millisecond()
    ))
}

/// Days since 1970-01-01 of the first of a month; `month` counts from 0 and
/// may lie outside the year, as in `Date.prototype.setMonth`.
fn month_start(year: i64, month: i64) -> Option<i64> {
    let year = year + month.div_euclid(12);
    let month = month.rem_euclid(12) + 1;
    let first = Date::new(i16::try_from(year).ok()?, month as i8, 1).ok()?;
    let epoch = Date::constant(1970, 1, 1);
    Some(i64::from((first - epoch).get_days()))
}

/// The `--since` value as `parseSince` in src/index.ts turns it into the
/// text the index is compared with: `30d`, `3m` and `1y` count back from now
/// on the local calendar, `YYYY-MM-DD` is that day's start in UTC, and
/// anything else is used as it stands.
pub fn since(value: &str, now: Timestamp) -> Result<String, HandBack> {
    since_in(value, now, local_zone)
}

fn since_in(value: &str, now: Timestamp, zone: impl FnOnce() -> Result<TimeZone, HandBack>) -> Result<String, HandBack> {
    let relative = value
        .char_indices()
        .last()
        .filter(|(at, unit)| *at > 0 && matches!(unit.to_ascii_lowercase(), 'd' | 'm' | 'y'))
        .filter(|(at, _)| value[..*at].bytes().all(|b| b.is_ascii_digit()));
    if let Some((at, unit)) = relative {
        // More than four digits is centuries; `Date` has answers there too
        // (years before 0, an invalid date), none worth copying.
        if at > 4 {
            return Err(HandBack);
        }
        let count: i64 = value[..at].parse().map_err(|_| HandBack)?;
        let zone = zone()?;
        let local = zone.to_datetime(now);
        // setDate / setMonth / setFullYear: the field is changed and the
        // date put together again, a day past the end of a month running
        // over into the next one.
        let (mut year, mut month, mut day) = (i64::from(local.year()), i64::from(local.month()) - 1, i64::from(local.day()));
        match unit.to_ascii_lowercase() {
            'd' => day -= count,
            'm' => month -= count,
            _ => year -= count,
        }
        let days = month_start(year, month).ok_or(HandBack)? + day - 1;
        let date = Date::constant(1970, 1, 1).checked_add(jiff::Span::new().try_days(days).map_err(|_| HandBack)?).map_err(|_| HandBack)?;
        let wall = DateTime::from_parts(date, Time::new(local.hour(), local.minute(), local.second(), local.subsec_nanosecond()).map_err(|_| HandBack)?);
        // A wall-clock time that a clock change skipped or repeated is read
        // with the offset from before the change, by both.
        let instant = zone.to_ambiguous_timestamp(wall).compatible().map_err(|_| HandBack)?;
        return iso_string(instant);
    }
    if value.len() == 10 && value.is_ascii() && value.as_bytes()[4] == b'-' && value.as_bytes()[7] == b'-' {
        let all_digits = [&value[0..4], &value[5..7], &value[8..10]].iter().all(|part| part.bytes().all(|b| b.is_ascii_digit()));
        if all_digits {
            // /^\d{4}-\d{2}-\d{2}$/
            let day = calendar_day(value).ok_or(HandBack)?;
            return Ok(format!("{:04}-{:02}-{:02}T00:00:00.000Z", day.year(), day.month(), day.day()));
        }
    }
    Ok(value.to_string())
}

/// `new Date(iso).toISOString().slice(0, 10)`: the UTC day of a time written
/// as `YYYY-MM-DDTHH:MM:SS[.fraction]` followed by `Z` or `±HH:MM`, which is
/// how Microsoft Graph writes them. `Date` reads much more than that.
pub fn utc_day(iso: &str) -> Result<String, HandBack> {
    let parse = || -> Option<String> {
        if !iso.is_ascii() || iso.len() < 20 {
            return None;
        }
        let day = calendar_day(&iso[0..10])?;
        let bytes = iso.as_bytes();
        if bytes[10] != b'T' || bytes[13] != b':' || bytes[16] != b':' {
            return None;
        }
        let (hour, minute, second) = (digits(&iso[11..13], 2)?, digits(&iso[14..16], 2)?, digits(&iso[17..19], 2)?);
        if hour > 23 || minute > 59 || second > 59 {
            return None;
        }
        let mut rest = &iso[19..];
        if let Some(fraction) = rest.strip_prefix('.') {
            let count = fraction.bytes().take_while(u8::is_ascii_digit).count();
            if count == 0 {
                return None;
            }
            rest = &fraction[count..];
        }
        let offset_seconds = match rest.as_bytes() {
            [b'Z'] => 0,
            [sign @ (b'+' | b'-'), _, _, b':', _, _] => {
                let (hours, minutes) = (digits(&rest[1..3], 2)?, digits(&rest[4..6], 2)?);
                if hours > 23 || minutes > 59 {
                    return None;
                }
                let seconds = hours * 3600 + minutes * 60;
                if *sign == b'+' {
                    seconds
                } else {
                    -seconds
                }
            }
            _ => return None,
        };
        let time = Time::new(hour as i8, minute as i8, second as i8, 0).ok()?;
        let wall = TimeZone::UTC.to_ambiguous_timestamp(DateTime::from_parts(day, time)).compatible().ok()?;
        let instant = Timestamp::from_second(wall.as_second() - offset_seconds).ok()?;
        let utc = TimeZone::UTC.to_datetime(instant);
        (1000..=9999).contains(&utc.year()).then(|| format!("{:04}-{:02}-{:02}", utc.year(), utc.month(), utc.day()))
    };
    parse().ok_or(HandBack)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(iso: &str) -> Timestamp {
        iso.parse().unwrap()
    }

    fn zoned(name: &str) -> TimeZone {
        zone(name).unwrap()
    }

    // Expected values were printed by bun 1.4.0 (`toLocaleString()` under
    // the same TZ, `new Date(text).toISOString()`).
    #[test]
    fn local_time_reads_like_en_us() {
        let t = at("2026-03-08T09:05:07Z").as_second();
        assert_eq!(locale_string(t, &zoned("UTC")).unwrap(), "3/8/2026, 9:05:07 AM");
        assert_eq!(locale_string(t, &zoned("Asia/Tokyo")).unwrap(), "3/8/2026, 6:05:07 PM");
        assert_eq!(locale_string(t, &zoned("America/New_York")).unwrap(), "3/8/2026, 5:05:07 AM");
        assert_eq!(locale_string(t, &zoned("US/Pacific")).unwrap(), "3/8/2026, 1:05:07 AM");
        assert_eq!(locale_string(t, &zoned("Etc/GMT+5")).unwrap(), "3/8/2026, 4:05:07 AM");
        assert_eq!(locale_string(at("2026-12-31T00:00:00Z").as_second(), &zoned("UTC")).unwrap(), "12/31/2026, 12:00:00 AM");
        assert_eq!(locale_string(at("2026-12-31T12:00:59Z").as_second(), &zoned("UTC")).unwrap(), "12/31/2026, 12:00:59 PM");
        assert!(zone("asia/tokyo").is_none());
        assert!(zone("JST-9").is_none());
        assert!(zone("").is_none());
    }

    #[test]
    fn a_page_shows_its_utc_day() {
        assert_eq!(utc_day("2026-03-01T10:00:00Z").unwrap(), "2026-03-01");
        assert_eq!(utc_day("2026-03-01T23:59:59.9999999Z").unwrap(), "2026-03-01");
        assert_eq!(utc_day("2026-03-01T10:00:00.1Z").unwrap(), "2026-03-01");
        assert_eq!(utc_day("2026-03-01T23:30:00-05:00").unwrap(), "2026-03-02");
        assert_eq!(utc_day("2026-03-01T00:30:00+09:00").unwrap(), "2026-02-28");
        assert_eq!(utc_day("2024-02-29T12:00:00Z").unwrap(), "2024-02-29");
        // Read by `Date` too, each in its own way; not taken on.
        for other in [
            "2026-03-01",
            "2026-03-01T10:00:00",
            "2026-03-01T10:00Z",
            "2026-02-30T00:00:00Z",
            "2026-13-01T00:00:00Z",
            "2026-03-01T24:00:00Z",
            "2026-03-01 10:00:00Z",
            "2026-03-01t10:00:00z",
            "March 1, 2026",
            "2026-03-01T10:00:00+0900",
            "2026-03-01T10:00:00.Z",
            "2026-03-01T10:00:60Z",
            " 2026-03-01T10:00:00Z",
            "2026-03-01T10:00:00+24:00",
            "",
        ] {
            assert!(utc_day(other).is_err(), "{other}");
        }
    }

    #[test]
    fn since_is_a_day_or_a_span_or_itself() {
        let now = at("2026-11-27T10:20:30.456Z");
        assert_eq!(since("2026-01-31", now).unwrap(), "2026-01-31T00:00:00.000Z");
        assert!(since("2026-02-30", now).is_err());
        assert!(since("2026-13-01", now).is_err());
        assert_eq!(since("2026-03-01T00:00:00Z", now).unwrap(), "2026-03-01T00:00:00Z");
        assert_eq!(since("yesterday", now).unwrap(), "yesterday");
        assert_eq!(since("d", now).unwrap(), "d");
        assert_eq!(since("3w", now).unwrap(), "3w");
        assert!(since("99999d", now).is_err());
    }

    /// The same sums as `setDate`, `setMonth` and `setFullYear` (values
    /// printed by bun 1.4.0 with TZ=UTC and TZ=America/New_York).
    #[test]
    fn spans_count_back_on_the_local_calendar() {
        let utc = |value: &str, now: &str| since_in(value, at(now), || Ok(zoned("UTC"))).unwrap();
        let now = "2026-03-31T10:20:30.456Z";
        assert_eq!(utc("30d", now), "2026-03-01T10:20:30.456Z");
        assert_eq!(utc("0d", now), "2026-03-31T10:20:30.456Z");
        assert_eq!(utc("1m", now), "2026-03-03T10:20:30.456Z");
        assert_eq!(utc("13M", now), "2025-03-03T10:20:30.456Z");
        assert_eq!(utc("400D", now), "2025-02-24T10:20:30.456Z");
        assert_eq!(utc("2y", "2024-02-29T12:00:00Z"), "2022-03-01T12:00:00.000Z");
        let new_york = |value: &str, now: &str| since_in(value, at(now), || Ok(zoned("America/New_York"))).unwrap();
        // Across the change to summer time the wall-clock time is kept.
        assert_eq!(new_york("30d", now), "2026-03-01T11:20:30.456Z");
        assert_eq!(new_york("13M", now), "2025-03-03T11:20:30.456Z");
        // 02:30 on 2026-03-08 does not exist in New York.
        assert_eq!(new_york("1d", "2026-03-09T06:30:00Z"), "2026-03-08T07:30:00.000Z");
        // 01:30 on 2026-11-01 happens twice there.
        assert_eq!(new_york("1d", "2026-11-02T06:30:00Z"), "2026-11-01T05:30:00.000Z");
        // A zone that cannot be read is only a problem when one is needed.
        assert!(since_in("1d", at(now), || Err(HandBack)).is_err());
        assert_eq!(since_in("2026-01-31", at(now), || Err(HandBack)).unwrap(), "2026-01-31T00:00:00.000Z");
    }
}
