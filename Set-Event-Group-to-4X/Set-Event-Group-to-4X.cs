using System;
using System.Collections.Generic;
using System.Collections;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace ComputerHelpNI
{
    public class SetEventGroupto4X
    {
        public Vegas myVegas;

        public void Main(Vegas vegas)
        {
            myVegas = vegas;

            double originalEventLength;
            double newEventLength;

            Dictionary<Track, List<TrackEvent>> selectedTrackEventMap;
            Dictionary<Track, List<TrackEvent>> nonSelectedtrackEventMap;
            GetSelectedTrackEvents(out selectedTrackEventMap, out nonSelectedtrackEventMap);

            if (selectedTrackEventMap.Count != 1)
            {
                MessageBox.Show("Requires: One selected track with selected event(s)");
                return;
            }

            //Assuming there's only 1 selected track, then get the selected track like this:
            IEnumerator selectedTrackEnum = selectedTrackEventMap.Keys.GetEnumerator();
            selectedTrackEnum.MoveNext();
            Track selectedTrack = (Track)selectedTrackEnum.Current;

            if (selectedTrack == null)
            {
                MessageBox.Show(selectedTrackEventMap.Count + "");
                return;
            }

            //Prevent acting on the default selection (e.g. from now to the end of the track)
            bool moreThanThreeMinutesSelected = false;
            double milisecondThreshold = 180000d; // 3 minutes
            foreach (TrackEvent trackEvent in selectedTrackEventMap[selectedTrack])
            {
                if (trackEvent.Length.ToMilliseconds() > milisecondThreshold)
                {
                    moreThanThreeMinutesSelected = true;
                    break;
                }
            }
            if (moreThanThreeMinutesSelected)
            {
                DialogResult dialogResult = MessageBox.Show("More than 3 minutes of video are selected.\nAre you sure you want to change all of it?", "Large Selection", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    //Exit
                    return;
                }
            }

            //When working with multicamera tracks, some events will be hidden from view and thus not selected nor part of the event group.
            //Build a new list of track events based on the event start time(s) of the selectedTrackEventMap items.

            //Get a list of start times.
            List<Timecode> startTimes = new List<Timecode>();
            foreach (TrackEvent trackEvent in selectedTrackEventMap[selectedTrack])
			{
                if (!startTimes.Contains(trackEvent.Start)) 
                { 
                    startTimes.Add(trackEvent.Start); 
                }
			}

            //Add all track events to a new event list (to include unselected events at the same start time(s) as the selected events)
            List<TrackEvent> workingEvents = new List<TrackEvent>();
            foreach (Timecode startTime in startTimes)
			{
                foreach (TrackEvent trackEvent in GetTrackEventsAtTime(startTime))
                {
                    if (trackEvent.Start == startTime)
                    {
                        workingEvents.Add(trackEvent);
                    }
                }
            }


            //Speed up each event at specified start time
            foreach (TrackEvent trackEvent in workingEvents)
            {
                originalEventLength = 0d;
                newEventLength = 0d;

                //Get value for new length
                if (originalEventLength == 0d)
                {
                    originalEventLength = trackEvent.Length.ToMilliseconds();
                    newEventLength = originalEventLength / 4d;
                }

                //Set playback speed to 400%
                trackEvent.AdjustPlaybackRate(400, true);

                //Adjust event length
                trackEvent.Length = Timecode.FromMilliseconds(newEventLength);

                //Ripple all of the next events back to close gaps
                for (int i = trackEvent.Index; i < (trackEvent.Track.Events.Count - 1); i++)
                {
                    if (trackEvent.Track.Events[i + 1] != null)
                    {
                        //Slide this event left by the amount of length removed from the selected event (original - new = amount to move)
                        trackEvent.Track.Events[i + 1].Start = trackEvent.Track.Events[i + 1].Start - Timecode.FromMilliseconds(originalEventLength - newEventLength);
                    }
                }
            }
        }

        public void GetSelectedTrackEvents(out Dictionary<Track, List<TrackEvent>> selectedTrackEventMap, out Dictionary<Track, List<TrackEvent>> nonSelectedtrackEventMap)
        {
            selectedTrackEventMap = new Dictionary<Track, List<TrackEvent>>();
            nonSelectedtrackEventMap = new Dictionary<Track, List<TrackEvent>>();

            foreach (Track track in myVegas.Project.Tracks)
            {
                if (track.IsValid())
                {
                    foreach (TrackEvent trackEvent in track.Events)
                    {
                        if (trackEvent.IsValid() && trackEvent.Selected)
                        {
                            Dictionary<Track, List<TrackEvent>> mapToUse = null;
                            if (track.Selected)
                                mapToUse = selectedTrackEventMap;
                            else
                                mapToUse = nonSelectedtrackEventMap;

                            if (!mapToUse.ContainsKey(track))
                            {
                                mapToUse.Add(track, new List<TrackEvent>());
                            }

                            mapToUse[track].Add(trackEvent);
                        }
                    }
                }
            }
        }

        public List<TrackEvent> GetTrackEventsAtTime(Timecode startTime)
        {
            List<TrackEvent> trackEvents = new List<TrackEvent>();

            foreach (Track track in myVegas.Project.Tracks)
            {
                if (track.IsValid())
                {
                    foreach (TrackEvent trackEvent in track.Events)
                    {
                        if (trackEvent.Start == startTime)
                        {
                            trackEvents.Add(trackEvent);
                        }
                    }
                }
            }

            return trackEvents;
        }

    }
}

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        ComputerHelpNI.SetEventGroupto4X test = new ComputerHelpNI.SetEventGroupto4X();
        test.Main(vegas);
    }
}
