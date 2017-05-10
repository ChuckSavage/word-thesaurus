/*
 * Made lots of changes from the original that I forked from. Thanks to Duncan Beaton for it.
 * 
 * Changes:
 * If there is no internet, displays "_ NO INTERNET _" in the menu.
 * If a search returns no results, doesn't display the menu
 * Plural searches return no results, remove trailing 's' and try again
 * Join the results, previously the results with the most values were kept, but it wasn't all of them
 * Keep a list of results for each search
 * Order the results alphabetically
 * Simplify URI creation, since it is the same for every search except for the word being looked up
 * Use Range instead of Selection for determining query
 * - Because, if the word right-clicked is not selected first, we expand to the word
 * - If the selection was used, it would expand it in the editor as well (it wasn't desirable)
 * Added an example of the XML returned from web-service to the end of this file
 * 
 * Bug fix: Duncan had included his uid in his code, removed it
 * 
 * Included my OrderedList collection, available at https://github.com/ChuckSavage/OrderedList
 * - It sorts additions, and in this implementation, ignores duplicates
 * 
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Xml.Linq;
using SeaRisenLib2.Collections;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace word_thesaurus
{
	public partial class ThisAddIn
	{
		/* The UID and KEY are in the file ThisAddIn_Credentials to make it easier to 
		 * upload changes to GitHub and not upload my keys
		 */

		const string NO_INTERNET = "_ No Internet _";

		private Word.Application app;
		//private Office.CommandBar commandBar;
		private Office.CommandBarPopup popup;
		Word.Range range;
		ArrayList menus = new ArrayList();
		List<int> temp = new List<int>();
		List<Office.CommandBar> textBars = new List<Office.CommandBar>();

		UriBuilder uri;
		string uritokens;
		SortedList lookup = new SortedList();
		OrderedList<string> words = new OrderedList<string>();

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			app = this.Application;
			app.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(app_WindowBeforeRightClick);
			// Get all bars, order by name for developer reference
			//words.Add(app.CommandBars.Cast<Office.CommandBar>()
			//	.Select(c => c.Name));

			bool txtadded = false;
			foreach (var bar in app.CommandBars)
			{
				var cmd = (Office.CommandBar)bar;
				switch (cmd.Name)
				{
					case "Text":           // Normal context-menu
					case "Spelling":       // Spelling context-menu
					case "Track Changes":  // Track changes context-menu
						if (cmd.Name == "Text") // There are two Text bars
						{                       // We only need one, using the first
							if (txtadded)
								continue;
							txtadded = true;
						}
						textBars.Add(cmd);
						break;
					default:
						continue;
				}
				//System.Diagnostics.Debug.WriteLine(cmd.Name);
			}

			uri = new UriBuilder();
			uri.Scheme = "http";
			uri.Host = "www.stands4.com";
			uri.Path = "services/v2/syno.php";
			uritokens = "uid=" + uid + "&tokenid=" + key + "&word=";
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			app.WindowBeforeRightClick -= new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(app_WindowBeforeRightClick);
			// these seem to always throw an exception
			//RemovePopup();
			//commandBar.Delete();
		}

		public void app_WindowBeforeRightClick(Word.Selection selection, ref bool Cancel)
		{
			if (null == selection) return;
			range = selection.Range;

			switch (selection.Type)
			{
				case Word.WdSelectionType.wdSelectionIP:
					range.Expand(Word.WdUnits.wdWord); // expand range to be a whole word                    
					break;
				case Word.WdSelectionType.wdSelectionNormal: // word(s) selected
					break;
				default: return;
			}
			// Trim range, so when we replace the query we don't trample on the spaces between words
			while (range.Characters.First.Text == " ")
				range.MoveStart(Word.WdUnits.wdCharacter);
			while (range.Characters.Last.Text == " ")
				range.MoveEnd(Word.WdUnits.wdCharacter, -1);

			if (!string.IsNullOrEmpty(range.Text))
				ShowMenu(range.Text);
		}

		public string UriSetup(string query)
		{
			uri.Query = uritokens + query;
			return uri.ToString();
		}

		public string[] Request(string query)
		{
			// If we've already done a search on this query, return it
			words = (OrderedList<string>)lookup[query];
			if (null != words)
				return words.ToArray();

			XElement xe;
			try
			{
				xe = XElement.Load(UriSetup(query));
			}
			catch (WebException e)
			{
				// if no network connection fail out nicely
				if (e.Message.StartsWith("The remote name could not be resolved"))
					return new string[] { NO_INTERNET }; // display message
				throw;
			}
			words = new OrderedList<string>();
			// See comment at end of file for XML result returned by web service
			// Add synonyms to words list, sorting and ignoring duplicates
			words.Add(
				xe.Descendants("synonyms")
				.SelectMany(xs => xs.Value.Split(','))
				.Select(s => s.Trim())
				);

			// Webservice doesn't fix plural words, and will return no results
			// so remove trailing 's' and retry
			if (words.Count < 1 && query.Length > 1 && query.ToLowerInvariant().EndsWith("s"))
			{
				// return is empty because of pluralism, rerun query without the trailing s
				// this will have an atrifact of the synonyms returned not being plural
				Request(query.Substring(0, query.Length - 1));
			}
			lookup.Add(query, words); // keep our sorted list for future look ups of same query
			return words.ToArray();
		}

		public void RemoveMenu()
		{
			foreach (Office.CommandBarButton button in menus)
			{
				button.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(button_Click);
				button.Visible = false;
				button.Delete();
			}
			menus.Clear();
			textBars.ForEach(b => b.Reset());
			//commandBar.Reset();
		}

		public void AddMenu(string[] words)
		{
			bool addmenu = true;
			// If no synonyms don't show the menu
			addmenu &= words.Length >= 1;
			// A search that returns nothing will have one empty result
			addmenu &= !(words.Length == 1 && string.IsNullOrWhiteSpace(words[0]));

			foreach (var commandBar in textBars)
			{
				//RemovePopup();
				popup = (Office.CommandBarPopup)commandBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, 1, false); // the 1 here is where in context menu the popup will appear, apparently 1 is min value
				popup.accName = "Theasurus";
				popup.Tag = "Theasurus";
				popup.Visible = addmenu;

				foreach (string word in words)
				{
					if (string.IsNullOrWhiteSpace(word)) continue;
					string w = word.Trim();
					var button = (Office.CommandBarButton)popup.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, popup.Controls.Count + 1, false);
					button.Caption = w;
					button.Tag = w;
					button.Visible = true;
					button.Click += new Office._CommandBarButtonEvents_ClickEventHandler(button_Click);
					menus.Add(button);
				}
			}
		}

		public void ShowMenu(string query)
		{
			//if (menus.Count > 0)
			RemoveMenu();
			string[] words = Request(query.ToLowerInvariant());
			AddMenu(words);
		}

		private void button_Click(Office.CommandBarButton ctrl, ref bool cancel)
		{
			if (ctrl.Tag != NO_INTERNET)
				range.Text = ctrl.Tag;
		}

		void RemovePopup()
		{
			try
			{
				if (null != popup)
					popup.Delete();
			}
			catch { }
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
/* 
 * XML return for "test"
 * 
<results>
  <result>
    <term>trial, trial run, test, tryout</term>
    <definition>trying something to find out about it</definition>
    <example>"a sample for ten days free trial"; "a trial of progesterone failed to relieve the pain"</example>
    <partofspeech>noun</partofspeech>
    <synonyms>test, mental test, tryout, trial, tribulation, exam, audition, psychometric test, mental testing, visitation, examination, trial run, run</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test, mental test, mental testing, psychometric test</term>
    <definition>any standardized procedure for measuring sensitivity or memory or intelligence or aptitude or personality etc</definition>
    <example>"the test was standardized on a large sample of students"</example>
    <partofspeech>noun</partofspeech>
    <synonyms>test, examination, tryout, trial, mental test, exam, psychometric test, mental testing, trial run, run</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>examination, exam, test</term>
    <definition>a set of questions or exercises evaluating skill or knowledge</definition>
    <example>"when the test was stolen the professor had to make a new set of questions"</example>
    <partofspeech>noun</partofspeech>
    <synonyms>scrutiny, examination, tryout, psychometric test, trial, mental test, exam, examen, interrogation, mental testing, test, run, interrogatory, trial run, testing</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test, trial</term>
    <definition>the act of undergoing testing</definition>
    <example>"he survived the great test of battle"; "candidates must compete in a trial of skill"</example>
    <partofspeech>noun</partofspeech>
    <synonyms>test, examination, tryout, trial, mental test, exam, psychometric test, mental testing, visitation, trial run, run, tribulation</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test, trial, run</term>
    <definition>the act of testing something</definition>
    <example>"in the experimental trials the amount of carbon was measured separately"; "he called each flip of the coin a new trial"</example>
    <partofspeech>noun</partofspeech>
    <synonyms>exam, running game, ladder, ravel, examination, political campaign, rill, running play, outpouring, tryout, tally, test, discharge, campaign, footrace, foot race, trial run, run, visitation, tribulation, streak, trial, psychometric test, rivulet, mental test, runnel, streamlet, mental testing, running</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test</term>
    <definition>a hard outer covering as of some amoebas and sea urchins</definition>
    <example></example>
    <partofspeech>verb</partofspeech>
    <synonyms>examination, tryout, trial, mental test, exam, psychometric test, mental testing, trial run, run</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test, prove, try, try out, examine, essay</term>
    <definition>put to the test, as for its quality, or give experimental use to</definition>
    <example>"This approach has been tried with good results"; "Test this recipe"</example>
    <partofspeech>verb</partofspeech>
    <synonyms>audition, rise, analyze, leaven, try, sample, show, testify, establish, test, evidence, try out, study, examine, see, seek, try on, probe, quiz, attempt, essay, raise, adjudicate, render, prove, shew, judge, screen, taste, turn out, demonstrate, turn up, experiment, strain, analyse, stress, hear, canvass, assay, canvas, bear witness</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>screen, test</term>
    <definition>test or examine for the presence of disease or infection</definition>
    <example>"screen the blood for the HIV virus"</example>
    <partofspeech>verb</partofspeech>
    <synonyms>test, examine, sieve, quiz, riddle, essay, try, prove, try out, screen, block out, sort, shield, screen out</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>quiz, test</term>
    <definition>examine someone's knowledge of something</definition>
    <example>"The teacher tests us every week"; "We got quizzed on French irregular verbs"</example>
    <partofspeech>verb</partofspeech>
    <synonyms>test, examine, quiz, essay, try, try out, screen, prove</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test</term>
    <definition>show a certain characteristic when tested</definition>
    <example>"He tested positive for HIV"</example>
    <partofspeech>verb</partofspeech>
    <synonyms>try, screen, essay, quiz, examine, prove, try out</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test</term>
    <definition>achieve a certain score or rating on a test</definition>
    <example>"She tested high on the LSAT and was admitted to all the good law schools"</example>
    <partofspeech>verb</partofspeech>
    <synonyms>try, screen, essay, quiz, examine, prove, try out</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test</term>
    <definition>determine the presence or properties of (a substance)</definition>
    <example></example>
    <partofspeech>verb</partofspeech>
    <synonyms>try, screen, essay, quiz, examine, prove, try out</synonyms>
    <antonyms></antonyms>
  </result>
  <result>
    <term>test</term>
    <definition>undergo a test</definition>
    <example>"She doesn't test well"</example>
    <partofspeech>verb</partofspeech>
    <synonyms>try, screen, essay, quiz, examine, prove, try out</synonyms>
    <antonyms></antonyms>
  </result>
</results>
*/
