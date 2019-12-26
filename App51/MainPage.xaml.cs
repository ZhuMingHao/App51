using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Media.Capture;
using Windows.Media.SpeechRecognition;
using Windows.Media.SpeechSynthesis;
using Windows.UI.Core;
using Windows.UI.Text;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace App51
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private const string ON = "enabled";
        private const string OFF = "disabled";
        private SpeechRecognizer speechRecognizer = new SpeechRecognizer(SpeechRecognizer.SystemSpeechLanguage);
        private CoreDispatcher dispatcher;
        private StringBuilder dictateBuilder = new StringBuilder();
        bool isListening = false;
        public MainPage()
        {
            this.InitializeComponent();

        }

        private void Cotent_TextChanged(object sender, RoutedEventArgs e)
        {
            richEbitBox.Document.GetText(TextGetOptions.None, out string value);
            charactersCount.Text = $"Characters: {value.Length - 1}";
        }
        protected async override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);
            bool permissionGained = await AudioCapturePermissions.RequestMicrophonePermission();
            if (permissionGained)
            {

                dispatcher = CoreWindow.GetForCurrentThread().Dispatcher;
                var dictationConstraint = new SpeechRecognitionTopicConstraint(SpeechRecognitionScenario.Dictation, "dictation");
                speechRecognizer.StateChanged += SpeechRecognizer_StateChanged;
                speechRecognizer.Constraints.Add(dictationConstraint);
                SpeechRecognitionCompilationResult result = await speechRecognizer.CompileConstraintsAsync();
                if (result.Status != SpeechRecognitionResultStatus.Success)
                {
                    return;
                }

                speechRecognizer.ContinuousRecognitionSession.Completed += ContinuousRecognitionSession_Completed;
                speechRecognizer.ContinuousRecognitionSession.ResultGenerated += ContinuousRecognitionSession_ResultGenerated;
                speechRecognizer.HypothesisGenerated += SpeechRecognizer_HypothesisGenerated;

            }
            else
            {

            }
        }

        private void SpeechRecognizer_StateChanged(SpeechRecognizer sender, SpeechRecognizerStateChangedEventArgs args)
        {

        }

        private async void Actions_Click(object sender, RoutedEventArgs e)
        {

            richEbitBox.Document.Selection.SetRange(0, richEbitBox.Document.Selection.EndPosition);

            var id = sender as Button;

            richEbitBox.Focus(FocusState.Pointer);

            switch (id.Tag)
            {

                case "0":
                    if (isListening == false)
                    {
                        // The recognizer can only start listening in a continuous fashion if the recognizer is currently idle.
                        // This prevents an exception from occurring.
                        if (speechRecognizer.State == SpeechRecognizerState.Idle)
                        {
                            isListening = true;
                            await speechRecognizer.ContinuousRecognitionSession.StartAsync();
                            isListening = false;
                        }
                    }
                    else
                    {
                        isListening = false;
                        if (speechRecognizer.State != SpeechRecognizerState.Idle)
                        {
                            // Cancelling recognition prevents any currently recognized speech from
                            // generating a ResultGenerated event. StopAsync() will allow the final session to
                            // complete.
                            await speechRecognizer.ContinuousRecognitionSession.StopAsync();
                        }
                    }
                    break;
                case "1":
                    if (richEbitBox.Document.Selection.CharacterFormat.Bold == FormatEffect.On)
                    {
                        richEbitBox.Document.Selection.CharacterFormat.Bold = FormatEffect.Off;
                        FormatBoltText.Background = (SolidColorBrush)Resources[OFF];
                    }
                    else
                    {
                        richEbitBox.Document.Selection.CharacterFormat.Bold = FormatEffect.On;
                        FormatBoltText.Background = (SolidColorBrush)Resources[ON];
                    }
                    break;
                case "2":
                    if (richEbitBox.Document.Selection.CharacterFormat.Italic == FormatEffect.On)
                    {
                        richEbitBox.Document.Selection.CharacterFormat.Italic = FormatEffect.Off;
                        formatItalicText.Background = (SolidColorBrush)Resources[OFF];
                    }
                    else
                    {
                        richEbitBox.Document.Selection.CharacterFormat.Italic = FormatEffect.On;
                        formatItalicText.Background = (SolidColorBrush)Resources[ON];
                    }
                    break;
                case "3":
                    if (richEbitBox.Document.Selection.CharacterFormat.Underline == UnderlineType.Single)
                    {
                        richEbitBox.Document.Selection.CharacterFormat.Underline = UnderlineType.None;
                        formatUnderlineText.Background = (SolidColorBrush)Resources[OFF];
                    }
                    else
                    {
                        richEbitBox.Document.Selection.CharacterFormat.Underline = UnderlineType.Single;
                        formatUnderlineText.Background = (SolidColorBrush)Resources[ON];
                    }
                    break;
                case "5":
                    richEbitBox.Document.GetText(TextGetOptions.AdjustCrlf, out string value);
                    speak(value);
                    break;
                default:
                    break;
            }
        }

        private async void SpeechRecognizer_HypothesisGenerated(
            SpeechRecognizer sender,
            SpeechRecognitionHypothesisGeneratedEventArgs args)
        {

            string hypothesis = args.Hypothesis.Text;
            string textboxContent = dictateBuilder.ToString() + " " + hypothesis + " ...";

            await dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                richEbitBox.Document.SetText(TextSetOptions.None, textboxContent);
            });
        }

        private async void ContinuousRecognitionSession_ResultGenerated(
            SpeechContinuousRecognitionSession sender,
            SpeechContinuousRecognitionResultGeneratedEventArgs args)
        {

            if (args.Result.Confidence == SpeechRecognitionConfidence.Medium ||
                  args.Result.Confidence == SpeechRecognitionConfidence.High)
            {

                dictateBuilder.Append(args.Result.Text + " ");

                await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
                {
                    richEbitBox.Document.SetText(TextSetOptions.None, dictateBuilder.ToString());
                });
            }
        }

        private void ContinuousRecognitionSession_Completed(
            SpeechContinuousRecognitionSession sender,
            SpeechContinuousRecognitionCompletedEventArgs args)
        {
        }
        private async void speak(string value)
        {

            MediaElement mediaElement = new MediaElement();

            var synth = new SpeechSynthesizer();

            VoiceInformation voiceInfo = (from voice in SpeechSynthesizer.AllVoices
                                          where voice.Gender == VoiceGender.Female
                                          select voice).FirstOrDefault();

            synth.Voice = voiceInfo;

            // Generate the audio stream from plain text.
            SpeechSynthesisStream stream = await synth.SynthesizeTextToStreamAsync(value);

            // Send the stream to the media object.
            mediaElement.SetSource(stream, stream.ContentType);
            mediaElement.Play();
        }

        private void ComboChanged(object sender, SelectionChangedEventArgs e)
        {

            richEbitBox.Focus(FocusState.Pointer);

            var id = sender as ComboBox;
            switch (id.Tag)
            {

                case "1":
                    string fontName = id.SelectedItem.ToString();
                    richEbitBox.Document.Selection.CharacterFormat.Name = fontName;
                    break;
                case "2":
                    var size = id.SelectedItem.ToString();
                    //set size to the Selection
                    richEbitBox.Document.Selection.CharacterFormat.Size = Convert.ToInt32(size);
                    break;
                default:
                    break;
            }
        }

        private void fontBox_Loaded(object sender, RoutedEventArgs e)
        {
            fontBox.Text = richEbitBox.FontFamily.Source.ToString();
            fontSizeBox.Text = richEbitBox.FontSize.ToString();
        }

        private void fontBox_TextSubmitted(ComboBox sender, ComboBoxTextSubmittedEventArgs args)
        {
            richEbitBox.Focus(FocusState.Pointer);
        }
    }

    public class AudioCapturePermissions
    {
        // If no recording device is attached, attempting to get access to audio capture devices will throw 
        // a System.Exception object, with this HResult set.
        private static int NoCaptureDevicesHResult = -1072845856;

        /// <summary>
        /// On desktop/tablet systems, users are prompted to give permission to use capture devices on a 
        /// per-app basis. Along with declaring the microphone DeviceCapability in the package manifest,
        /// this method tests the privacy setting for microphone access for this application.
        /// Note that this only checks the Settings->Privacy->Microphone setting, it does not handle
        /// the Cortana/Dictation privacy check, however (Under Settings->Privacy->Speech, Inking and Typing).
        /// 
        /// Developers should ideally perform a check like this every time their app gains focus, in order to 
        /// check if the user has changed the setting while the app was suspended or not in focus.
        /// </summary>
        /// <returns>true if the microphone can be accessed without any permissions problems.</returns>
        public async static Task<bool> RequestMicrophonePermission()
        {
            try
            {
                // Request access to the microphone only, to limit the number of capabilities we need
                // to request in the package manifest.
                MediaCaptureInitializationSettings settings = new MediaCaptureInitializationSettings();
                settings.StreamingCaptureMode = StreamingCaptureMode.Audio;
                settings.MediaCategory = MediaCategory.Speech;
                MediaCapture capture = new MediaCapture();

                await capture.InitializeAsync(settings);
            }
            catch (TypeLoadException)
            {
                // On SKUs without media player (eg, the N SKUs), we may not have access to the Windows.Media.Capture
                // namespace unless the media player pack is installed. Handle this gracefully.
                var messageDialog = new Windows.UI.Popups.MessageDialog("Media player components are unavailable.");
                await messageDialog.ShowAsync();
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                // The user has turned off access to the microphone. If this occurs, we should show an error, or disable
                // functionality within the app to ensure that further exceptions aren't generated when 
                // recognition is attempted.
                return false;
            }
            catch (Exception exception)
            {
                // This can be replicated by using remote desktop to a system, but not redirecting the microphone input.
                // Can also occur if using the virtual machine console tool to access a VM instead of using remote desktop.
                if (exception.HResult == NoCaptureDevicesHResult)
                {
                    var messageDialog = new Windows.UI.Popups.MessageDialog("No Audio Capture devices are present on this system.");
                    await messageDialog.ShowAsync();
                    return false;
                }
                else
                {
                    throw;
                }
            }
            return true;
        }
    }
}
