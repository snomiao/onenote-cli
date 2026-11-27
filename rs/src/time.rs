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

/// Where the JavaScript runtime reads a zone differently from this binary:
/// the zone, and from when until when.
///
/// The runtime carries its own time zone database (ICU's, inside bun) and
/// this binary carries one too (the `jiff-tzdb` crate), and the two are not
/// of the same edition. Where a country has changed its clocks between the
/// editions, or an old change was corrected, they are an hour apart for a
/// stretch of time. Listed here is what differs between bun 1.4.0 (ICU 78.3)
/// and tzdb 2026f; the binary leaves these stretches to the runtime.
///
/// The test `every_zone_is_read_as_the_runtime_reads_it` at the end of this
/// file compares every zone with the runtime that is installed, fails for a
/// difference that is not listed, and prints the line to add.
const READ_DIFFERENTLY_BY_THE_RUNTIME: [(&str, &str, &str); 16] = [
    // Morocco, from September 2026 on.
    ("Africa/Casablanca", "2026-09-20T01:00:00Z", "2099-12-31T10:05:09Z"),
    ("Africa/El_Aaiun", "2026-09-20T01:00:00Z", "2099-12-31T10:05:09Z"),
    // Colombia: the day in May 1992 on which summer time began.
    ("America/Bogota", "1992-05-02T05:00:00Z", "1992-05-02T13:52:59Z"),
    // Western Canada: every winter from November 2026 on. The compiled-in
    // database keeps these zones on summer time, the runtime's does not.
    ("America/Edmonton", "2026-11-01T08:00:00Z", "2099-12-31T10:05:09Z"),
    ("America/Inuvik", "2026-11-01T08:00:00Z", "2099-12-31T10:05:09Z"),
    ("America/Rainy_River", "2026-11-01T07:00:00Z", "2099-12-31T10:05:09Z"),
    ("America/Vancouver", "2026-11-01T09:00:00Z", "2099-12-31T10:05:09Z"),
    ("America/Winnipeg", "2026-11-01T07:00:00Z", "2099-12-31T10:05:09Z"),
    ("America/Yellowknife", "2026-11-01T08:00:00Z", "2099-12-31T10:05:09Z"),
    ("Canada/Central", "2026-11-01T07:00:00Z", "2099-12-31T10:05:09Z"),
    ("Canada/Mountain", "2026-11-01T08:00:00Z", "2099-12-31T10:05:09Z"),
    ("Canada/Pacific", "2026-11-01T09:00:00Z", "2099-12-31T10:05:09Z"),
    // Cuba: the winter of 2026, an hour apart as well.
    ("America/Havana", "2026-11-01T15:19:55Z", "2027-03-13T00:31:27Z"),
    ("Cuba", "2026-11-01T15:19:55Z", "2027-03-13T00:31:27Z"),
    // Iran: an hour of one night in May 1979.
    ("Asia/Tehran", "1979-05-25T20:30:00Z", "1979-05-25T21:30:01Z"),
    ("Iran", "1979-05-25T20:30:00Z", "1979-05-25T21:30:01Z"),
];

/// How far the binary stays away from a listed stretch, on both sides. The
/// test steps through the years in strides shorter than this, so a stretch
/// can begin or end that much outside of what the test has seen; and a time
/// of day that comes twice is read with the clock from before the change,
/// which can lie on the other side of a stretch's edge.
const MARGIN: i64 = 3 * 86400;

/// The years over which that test compares the zones, 1970 to 2099, as
/// seconds since 1970. Outside them nothing is known about the two databases
/// (before 1970 they are known to differ: the runtime's keeps old clock
/// changes that the compiled-in one has dropped).
const COMPARED: std::ops::Range<i64> = 0..4_102_444_800;

/// A time zone of the database compiled into the binary, and the stretches
/// of time for which it is left to the runtime.
pub struct Zone {
    zone: TimeZone,
    apart: Vec<std::ops::Range<i64>>,
}

impl Zone {
    /// A zone by the name it has in the database. The runtime also takes
    /// `asia/tokyo`; that is not taken here.
    fn named(name: &str) -> Option<Zone> {
        let zone = jiff::tz::TimeZoneDatabase::bundled().get(name).ok()?;
        if zone.iana_name() != Some(name) {
            return None;
        }
        let second = |time: &str| time.parse::<Timestamp>().ok().map(Timestamp::as_second);
        let apart = READ_DIFFERENTLY_BY_THE_RUNTIME
            .iter()
            .filter(|(listed, ..)| *listed == name)
            .map(|(_, from, until)| Some(second(from)? - MARGIN..second(until)? + MARGIN))
            .collect::<Option<Vec<_>>>()?;
        Some(Zone { zone, apart })
    }

    /// `Ok` for a moment at which the binary and the runtime were seen to
    /// read this zone the same way.
    fn agreed(&self, instant: Timestamp) -> Result<(), HandBack> {
        let second = instant.as_second();
        if COMPARED.contains(&second) && !self.apart.iter().any(|stretch| stretch.contains(&second)) {
            Ok(())
        } else {
            Err(HandBack)
        }
    }
}

/// The zone `Date` uses for local time: the one `TZ` names, else the
/// system's. `TZ` can also hold a rule (`JST-9`), a path or nothing, which
/// the JavaScript runtime reads its own way.
pub fn local_zone() -> Result<Zone, HandBack> {
    match std::env::var_os("TZ") {
        Some(tz) => tz.to_str().and_then(Zone::named).ok_or(HandBack),
        None => TimeZone::try_system().ok().and_then(|system| system.iana_name().and_then(Zone::named)).ok_or(HandBack),
    }
}

/// The second a file's modification time falls in, as the runtime's `Date`
/// for it (`stat().mtime`) has it.
///
/// The runtime makes the time a number of milliseconds with a fraction
/// (seconds * 1000 + nanoseconds / 1e6, in floating point) and `Date` drops
/// the fraction. Close enough to the end of a second the sum rounds up to
/// the next one: bun 1.4.0 shows the next second from 122 ns before it (in
/// 2026). Rather than repeat that arithmetic, the last millisecond of every
/// second is left to the runtime, which also covers one that would round to
/// the nearest millisecond, as Node does.
pub fn modified_second(modified: std::time::SystemTime) -> Result<i64, HandBack> {
    let since_1970 = modified.duration_since(std::time::UNIX_EPOCH).map_err(|_| HandBack)?;
    if since_1970.subsec_nanos() >= 999_000_000 {
        return Err(HandBack);
    }
    i64::try_from(since_1970.as_secs()).map_err(|_| HandBack)
}

/// `new Date(seconds * 1000).toLocaleString()`. The runtime's locale is
/// en-US whatever the user's is (bun does not read LANG or LC_*).
pub fn locale_string(seconds: i64, zone: &Zone) -> Result<String, HandBack> {
    let instant = Timestamp::from_second(seconds).map_err(|_| HandBack)?;
    zone.agreed(instant)?;
    Ok(en_us(zone.zone.to_datetime(instant)))
}

/// A date and time as the en-US locale writes them: `3/8/2026, 9:05:07 AM`.
fn en_us(local: DateTime) -> String {
    let (hour, half) = match local.hour() {
        0 => (12, "AM"),
        hour @ 1..=11 => (hour, "AM"),
        12 => (12, "PM"),
        hour => (hour - 12, "PM"),
    };
    format!("{}/{}/{}, {}:{:02}:{:02} {}", local.month(), local.day(), local.year(), hour, local.minute(), local.second(), half)
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

fn since_in(value: &str, now: Timestamp, zone: impl FnOnce() -> Result<Zone, HandBack>) -> Result<String, HandBack> {
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
        zone.agreed(now)?;
        let local = zone.zone.to_datetime(now);
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
        let instant = zone.zone.to_ambiguous_timestamp(wall).compatible().map_err(|_| HandBack)?;
        zone.agreed(instant)?;
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

    fn zoned(name: &str) -> Zone {
        Zone::named(name).unwrap()
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
        assert!(Zone::named("asia/tokyo").is_none());
        assert!(Zone::named("JST-9").is_none());
        assert!(Zone::named("").is_none());
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

    #[test]
    fn only_the_years_that_were_compared_are_answered() {
        let utc = zoned("UTC");
        assert_eq!(locale_string(0, &utc).unwrap(), "1/1/1970, 12:00:00 AM");
        assert!(locale_string(-1, &utc).is_err());
        assert_eq!(locale_string(COMPARED.end - 1, &utc).unwrap(), "12/31/2099, 11:59:59 PM");
        assert!(locale_string(COMPARED.end, &utc).is_err());
        let since = |value: &str, now: &str| since_in(value, at(now), || Ok(zoned("UTC")));
        assert_eq!(since("56y", "2026-11-27T10:20:30.456Z").unwrap(), "1970-11-27T10:20:30.456Z");
        assert!(since("57y", "2026-11-27T10:20:30.456Z").is_err());
        assert!(since("1d", "2100-01-01T00:00:00Z").is_err());
        assert!(since("1d", "1969-12-31T23:59:59Z").is_err());
        // A day is no span: nothing of it depends on a zone.
        assert_eq!(since("1026-01-31", "2026-11-27T10:20:30.456Z").unwrap(), "1026-01-31T00:00:00.000Z");
    }

    #[test]
    fn a_listed_stretch_is_left_to_the_runtime_with_a_margin() {
        for (name, from, until) in READ_DIFFERENTLY_BY_THE_RUNTIME {
            let zone = Zone::named(name).unwrap_or_else(|| panic!("{name}: not a zone of the database, or a time that cannot be read"));
            let (from, until) = (at(from).as_second(), at(until).as_second());
            assert!(from < until, "{name}");
            for inside in [from - MARGIN, from, until - 1, until + MARGIN - 1] {
                assert!(locale_string(inside, &zone).is_err(), "{name} at {inside}");
            }
            for outside in [from - MARGIN - 1, until + MARGIN] {
                assert_eq!(locale_string(outside, &zone).is_ok(), COMPARED.contains(&outside), "{name} at {outside}");
            }
        }
    }

    #[test]
    fn the_last_millisecond_of_a_second_is_left_to_the_runtime() {
        let at = |seconds: u64, nanoseconds: u32| std::time::UNIX_EPOCH + std::time::Duration::new(seconds, nanoseconds);
        assert_eq!(modified_second(at(1_772_960_707, 0)).unwrap(), 1_772_960_707);
        assert_eq!(modified_second(at(1_772_960_707, 998_999_999)).unwrap(), 1_772_960_707);
        // bun 1.4.0 shows :07 for the first two of these and :08 for the third.
        assert!(modified_second(at(1_772_960_707, 999_000_000)).is_err());
        assert!(modified_second(at(1_772_960_707, 999_999_877)).is_err());
        assert!(modified_second(at(1_772_960_707, 999_999_878)).is_err());
        assert!(modified_second(std::time::UNIX_EPOCH - std::time::Duration::new(1, 0)).is_err());
    }

    /// Questions for rs/parity/runtime-time.ts about one zone, each with the
    /// answer this file gives.
    struct Questions<'z> {
        zone: &'z Zone,
        /// What the binary answers itself: the runtime has to say the same.
        asked: Vec<(String, String)>,
        /// What it leaves to the runtime because of the list. Asked as well,
        /// to see whether the list still holds.
        left: Vec<(String, String)>,
    }

    impl Questions<'_> {
        /// "last synced" for a file modified at that second.
        fn local(&mut self, seconds: i64) {
            match locale_string(seconds, self.zone) {
                Ok(answer) => self.asked.push((format!("U {seconds}"), answer)),
                Err(_) if COMPARED.contains(&seconds) => {
                    let answer = en_us(self.zone.zone.to_datetime(Timestamp::from_second(seconds).unwrap()));
                    self.left.push((format!("U {seconds}"), answer));
                }
                Err(_) => {}
            }
        }

        /// `--since <value>` at that moment.
        fn since(&mut self, now: Timestamp, value: &str) {
            let zone = Zone { zone: self.zone.zone.clone(), apart: self.zone.apart.clone() };
            if let Ok(answer) = since_in(value, now, || Ok(zone)) {
                self.asked.push((format!("S {} {value}", now.as_millisecond()), answer));
            }
        }

        /// The years compared in strides of a day and a bit (so that every
        /// hour of the day comes up), every clock change of the zone from
        /// both sides, and both ends of every listed stretch.
        fn all(zone: &Zone) -> Questions<'_> {
            let mut questions = Questions { zone, asked: Vec::new(), left: Vec::new() };
            let second = |seconds: i64| Timestamp::from_second(seconds).unwrap();
            let mut stride = 0;
            let mut at = COMPARED.start + 3 * 3600 + 17;
            while at < COMPARED.end {
                questions.local(at);
                if stride % 320 == 0 {
                    let now = Timestamp::from_millisecond(at * 1000 + 456).unwrap();
                    for value in ["30d", "3m", "13M", "10Y"] {
                        questions.since(now, value);
                    }
                }
                stride += 1;
                at += STRIDE;
            }
            for change in zone.zone.following(second(COMPARED.start)) {
                let at = change.timestamp().as_second();
                if at >= COMPARED.end {
                    break;
                }
                for around in [-3600, -1, 0, 1, 3600] {
                    questions.local(at + around);
                }
                // The times a clock showed around the change, by the offset
                // before it and by the one after it: among them are times it
                // skipped and times it showed twice. Each is what `--since`
                // has to come back to, from a day, a month and a year later.
                for offset in [zone.zone.to_offset(second(at - 1)), change.offset()] {
                    for around in [-1, 0, 1800, 3600] {
                        let shown = offset.to_datetime(second(at + around));
                        let mut back_to = |value: &str, span: jiff::Span| {
                            let Ok(later) = shown.checked_add(span) else { return };
                            let Ok(now) = zone.zone.to_ambiguous_timestamp(later).compatible() else { return };
                            let Ok(now) = Timestamp::from_millisecond(now.as_millisecond() + 456) else { return };
                            questions.since(now, value);
                        };
                        back_to("1d", jiff::Span::new().days(1));
                        if around == 1800 {
                            back_to("30D", jiff::Span::new().days(30));
                            back_to("1m", jiff::Span::new().months(1));
                            back_to("1y", jiff::Span::new().years(1));
                        }
                    }
                }
            }
            for stretch in &zone.apart {
                for outside in [stretch.start - 3600, stretch.start - 1, stretch.end, stretch.end + 3600] {
                    questions.local(outside);
                }
            }
            questions
        }
    }

    /// The stride of that walk through the years. Shorter than `MARGIN`, so
    /// that a stretch found by it, widened by the margin, is all of it.
    const STRIDE: i64 = 86400 + 7 * 3600 + 13 * 60 + 11;

    /// The answers of the JavaScript runtime, started the way `onenote` is:
    /// with the zone in TZ.
    fn ask_the_runtime(zone: &str, questions: &[&(String, String)], directory: &std::path::Path) -> Vec<String> {
        use std::io::Write;
        use std::process::{Command, Stdio};
        let script = concat!(env!("CARGO_MANIFEST_DIR"), "/parity/runtime-time.ts");
        let mut runtime = Command::new("bun")
            .arg(script)
            .current_dir(directory)
            .env_clear()
            .env("PATH", std::env::var_os("PATH").unwrap_or_default())
            .env("HOME", directory)
            .env("TZ", zone)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .spawn()
            .expect("this test needs bun on PATH");
        let asked: String = questions.iter().map(|(question, _)| format!("{question}\n")).collect();
        // The script reads every question before it answers the first.
        runtime.stdin.take().unwrap().write_all(asked.as_bytes()).unwrap();
        let output = runtime.wait_with_output().unwrap();
        assert!(output.status.success(), "bun failed for {zone}");
        let answers: Vec<String> = String::from_utf8(output.stdout).unwrap().lines().map(str::to_string).collect();
        assert_eq!(answers.len(), questions.len(), "{zone}: as many answers as questions");
        answers
    }

    fn iso(seconds: i64) -> String {
        Timestamp::from_second(seconds).unwrap().to_string()
    }

    /// What came of comparing one zone with the runtime.
    struct Found {
        name: String,
        compared: usize,
        spans: usize,
        /// Answers that differ, and the times each is about, in seconds:
        /// the one asked for, or for `--since` the moment it was asked at
        /// and the one it came back to.
        differing: Vec<(String, Vec<i64>)>,
        /// A listed zone on which the runtime differs nowhere any more.
        listed_for_nothing: bool,
    }

    fn compare_with_the_runtime(name: &str, directory: &std::path::Path) -> Found {
        let zone = Zone::named(name).unwrap_or_else(|| panic!("{name}: not read"));
        let questions = Questions::all(&zone);
        let every: Vec<&(String, String)> = questions.asked.iter().chain(&questions.left).collect();
        let answers = ask_the_runtime(name, &every, directory);
        let (asked, left) = answers.split_at(questions.asked.len());
        let differing = questions
            .asked
            .iter()
            .zip(asked)
            .filter(|((_, mine), theirs)| mine != *theirs)
            .map(|((question, mine), theirs)| {
                let asked: i64 = question.split(' ').nth(1).unwrap().parse().unwrap();
                let times = if question.starts_with('U') { vec![asked] } else { vec![asked / 1000, at(mine).as_second()] };
                (format!("`{question}`: {mine} here, {theirs} in the runtime"), times)
            })
            .collect();
        Found {
            name: name.to_string(),
            compared: questions.asked.len(),
            spans: questions.asked.iter().filter(|(question, _)| question.starts_with('S')).count(),
            differing,
            listed_for_nothing: !zone.apart.is_empty() && questions.left.iter().zip(left).all(|((_, mine), theirs)| mine == theirs),
        }
    }

    /// Every zone name in the database compiled into the binary, against the
    /// JavaScript runtime (bun, which has to be on PATH): "last synced" for
    /// every 31 hours from 1970 to 2099 and around every clock change, and
    /// `--since` spans that come back to every clock change. Where the two
    /// differ, the zone and the stretch of time have to be listed in
    /// `READ_DIFFERENTLY_BY_THE_RUNTIME`, so that the binary does not answer
    /// there.
    ///
    /// When this fails because the runtime or the `jiff-tzdb` crate has
    /// moved to another edition of the time zone database, put the lines it
    /// prints into the list. It takes a minute or two.
    #[test]
    fn every_zone_is_read_as_the_runtime_reads_it() {
        use std::sync::atomic::{AtomicUsize, Ordering};
        use std::sync::Mutex;
        let names: Vec<String> = jiff::tz::TimeZoneDatabase::bundled().available().map(|name| name.as_str().to_string()).collect();
        assert!(names.len() > 500, "the database has {} zones", names.len());
        // An empty directory to start the runtime in: it reads .env files
        // and bunfig.toml from where it is started.
        let directory = std::env::temp_dir().join(format!("onenote-cli-zones-{}", std::process::id()));
        std::fs::create_dir_all(&directory).unwrap();

        let next = AtomicUsize::new(0);
        let found: Mutex<Vec<Found>> = Mutex::new(Vec::new());
        std::thread::scope(|scope| {
            for _ in 0..std::thread::available_parallelism().map_or(4, |n| n.get()).min(8) {
                scope.spawn(|| {
                    while let Some(name) = names.get(next.fetch_add(1, Ordering::Relaxed)) {
                        let zone = compare_with_the_runtime(name, &directory);
                        found.lock().unwrap().push(zone);
                    }
                });
            }
        });
        let _ = std::fs::remove_dir_all(&directory);

        let mut found = found.into_inner().unwrap();
        found.sort_by(|a, b| a.name.cmp(&b.name));
        let compared: usize = found.iter().map(|zone| zone.compared).sum();
        let spans: usize = found.iter().map(|zone| zone.spans).sum();
        println!("{} zones, {compared} answers compared with the runtime, {spans} of them --since spans", found.len());
        assert!(found.len() == names.len() && compared > 10_000_000, "too little was compared: {} zones, {compared} answers", found.len());
        for zone in found.iter().filter(|zone| zone.listed_for_nothing) {
            println!("{} is listed as read differently, but this runtime agrees on it everywhere", zone.name);
        }
        let mut wrong = Vec::new();
        for zone in &found {
            // Local times that differ less than 400 days apart are one
            // stretch: a difference that comes back every winter is one line.
            let mut local: Vec<(i64, &str)> = zone.differing.iter().filter(|(_, times)| times.len() == 1).map(|(example, times)| (times[0], example.as_str())).collect();
            local.sort();
            let mut stretches: Vec<(i64, i64, usize, &str)> = Vec::new();
            for (seconds, example) in local {
                match stretches.last_mut() {
                    Some((_, until, count, _)) if seconds - *until <= 400 * 86400 => (*until, *count) = (seconds + 1, *count + 1),
                    _ => stretches.push((seconds, seconds + 1, 1, example)),
                }
            }
            for (from, until, count, example) in &stretches {
                wrong.push(format!("    (\"{}\", \"{}\", \"{}\"), // {count} local times, e.g. {example}", zone.name, iso(*from), iso(*until)));
            }
            // A `--since` that differs because it starts or ends in such a
            // stretch needs no line of its own.
            let explained = |times: &[i64]| times.iter().any(|time| stretches.iter().any(|(from, until, ..)| (from - MARGIN..until + MARGIN).contains(time)));
            for (example, _) in zone.differing.iter().filter(|(_, times)| times.len() == 2 && !explained(times)) {
                wrong.push(format!("    {}: a --since outside of those stretches: {example}", zone.name));
            }
        }
        assert!(wrong.is_empty(), "the runtime reads these differently, and they are not listed in READ_DIFFERENTLY_BY_THE_RUNTIME:\n{}", wrong.join("\n"));
    }
}
