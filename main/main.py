#!/usr/bin/python3.7

import os
import sys
import numpy as np
import crepe
import sox
import pandas as pd


def main(args):
    """
    ---------------------------------------------------
    Main method, gets called when the script executes.
    ---------------------------------------------------
    :param args: Input provided when the script is run.
                    args[1]: Source directory
                    args[2]: Destination directory
                    args[3]: Excel sheet name (extension
                             exclusive)
    ---------------------------------------------------
    """

    # Extract arguments from console
    from_dir = args[1]
    to_dir = args[2]
    excel_name = args[3]

    # Inform user of where files will be read/written
    print("Audio files will be read from:", from_dir)
    print("Pitch shifted files and excel sheet will be written to:", to_dir)
    print("Excel file will be named:", excel_name)

    # Verify if from_dir is valid directory
    if not os.path.isdir(from_dir):
        raise Exception(from_dir + " is not a valid directory. Terminating application.")

    # Verify if to_dir is valid directory
    if not os.path.isdir(to_dir):
        raise Exception(to_dir + " is not a valid directory. Terminating application.")

    # Start pitch shifting on source directory
    pitch_shift_file(from_dir, to_dir, excel_name)


def get_analysis_dict(speech_array, sample_rate):
    """
    ------------------------------------------------------
    Uses CREPE to analyze a provided speech array. Returns
    a dictionary containing the sample rate and the
    minimum, maximum, average and median frequency.
    ------------------------------------------------------
    :param speech_array: speech array to analyze
    :param sample_rate: sample rate of the speech segment
    :return: dictionary
    """
    analysis = dict()

    # CREPE Analysis
    time, frequencies, confidence, activation = crepe.predict(speech_array, sample_rate,
                                                              step_size=2500, viterbi=True)
    # Get frequency information
    min_frequency = np.min(frequencies)
    max_frequency = np.max(frequencies)
    med_frequency = np.median(frequencies)
    avg_frequency = np.average(frequencies)

    # Put information in dictionary
    analysis["min_freq"] = min_frequency
    analysis["max_freq"] = max_frequency
    analysis["med_freq"] = med_frequency
    analysis["avg_freq"] = avg_frequency
    analysis["sample_rate"] = sample_rate

    return analysis


def append_to_dataframe(dataframe, filename, ps_filename, semitone_difference,
                        new_pitch, o_analysis, ps_analysis, length):
    """
    ------------------------------------------------------------
    Appends a new row of pitch-shift data to a pandas dataframe.
    ------------------------------------------------------------
    :param dataframe: pandas dataframe to append to
    :param filename: file name of source speech
    :param ps_filename: file name of pitch-shifted speech
    :param semitone_difference: semitone difference between
                                the two speech files
    :param new_pitch: target pitch
    :param o_analysis: analysis dictionary of source speech
    :param ps_analysis: analysis dictionary of pitch-shifted
                        speech
    :param length: length of the file

    :return: the appended pandas dataframe
    """
    # Excel row for current file
    analysis = {"o_file": filename, "ps_file": ps_filename, "length": length, "sample_rate": o_analysis["sample_rate"],
                "target_freq": new_pitch, "o_med_freq": o_analysis["med_freq"], "ps_med_freq": ps_analysis["med_freq"],
                "semitone_difference": semitone_difference,
                "o_min_freq": o_analysis["min_freq"], "ps_min_freq": ps_analysis["min_freq"],
                "o_max_freq": o_analysis["max_freq"], "ps_max_freq": ps_analysis["max_freq"],
                "o_avg_freq": o_analysis["avg_freq"], "ps_avg_freq": ps_analysis["avg_freq"]}

    # Append row to dataframe and return
    return dataframe.append(analysis, ignore_index=True)


def pitch_shift_file(from_dir, to_dir, excel_name):
    """
    ------------------------------------------------------------------------------------
    Performs a pitch shift to the range [250,300) on the entire file at once. Uses CRÊPE
    to find the frequencies of the original file. New files are written to output
    directory as original_filename_ps_file.wav. An Excel file containing file
    information will be created in the output directory. This contains information such
    as file name, new file name, pitch-shift... .
    ------------------------------------------------------------------------------------
    :param from_dir: input directory
    :param to_dir: output directory
    :param excel_name: name for Excel sheet (extension exclusive)
    ------------------------------------------------------------------------------------
    """

    # Create headers for Excel file
    excel_headers = ["o_file", "ps_file", "length", "sample_rate", "target_freq", "o_med_freq",
                     "ps_med_freq", "semitone_difference", "o_min_freq", "ps_min_freq",
                     "o_max_freq", "ps_max_freq", "o_avg_freq", "ps_avg_freq"]

    # Pandas dataframe to keep track of file analysis
    dataframe = pd.DataFrame(columns=excel_headers)

    # Loop over all files in from directory
    for o_filename in os.listdir(from_dir):

        # Only execute on .wav files
        if o_filename.endswith(".wav"):

            print("Now shifting", o_filename)

            # Get file name and path
            stripped_filename = os.path.splitext(o_filename)[0]
            from_path = from_dir + "/" + o_filename

            # In-memory array source speech
            mem_tfm = sox.Transformer()
            o_speech_array = mem_tfm.build_array(from_path)
            sample_rate = np.floor(sox.file_info.sample_rate(from_path))

            # Analysis source speech
            o_analysis = get_analysis_dict(o_speech_array, sample_rate)
            o_med_frequency = o_analysis["med_freq"]

            # Target frequency
            new_pitch = np.random.randint(250, 300)

            semitone_difference = 12 * np.log2(new_pitch / o_med_frequency)

            ps_speech_array = []

            # Perform pitch shift
            try:
                # Trim to correct interval
                tfm = sox.Transformer()

                # Prevent clipping
                tfm.gain(limiter=True)

                # Stereo to mono
                if sox.file_info.channels(from_path) >= 2:
                    tfm.oops()

                # Perform pitch shift on entire file
                tfm.pitch(semitone_difference)

                # Prevent clipping
                tfm.gain(limiter=True)

                # Output shifted array
                ps_speech_array = tfm.build_array(sample_rate_in=sample_rate, input_array=o_speech_array)

            # Exception occurred when shifting
            except Exception as e:
                print("Could not do pitch shift for file", o_filename)
                print("An exception occurred:", e)

            print("Completed", o_filename)

            # Give correct file name if extreme shift pitch
            is_extreme = semitone_difference < -12 or semitone_difference > 12
            ps_filename = stripped_filename + "_pse.wav" if is_extreme else stripped_filename + "_ps.wav"
            to_path = to_dir + "/" + ps_filename

            # Write shifted file to target path
            mem_tfm.build_file(output_filepath=to_path,
                               sample_rate_in=sample_rate,
                               input_array=np.array(ps_speech_array))

            # Analyze result
            ps_analysis = get_analysis_dict(np.array(ps_speech_array), sample_rate)

            # Append to analysis dataframe
            dataframe = append_to_dataframe(dataframe, o_filename, ps_filename, semitone_difference,
                                            new_pitch, o_analysis, ps_analysis, sox.file_info.duration(from_path))

    # Write dataframe to Excel file
    write_excel_file(dataframe, excel_name, to_dir)

def write_excel_file(dataframe, excel_name, to_dir):
    """
    ------------------------------------------------------------------------------------
    Given a dataframe, creates the Excel file with a given name in a given directory.
    ------------------------------------------------------------------------------------
    :param dataframe: Dataframe to write.
    :param excel_name: Name of the Excel file.
    :param to_dir: Target directory.
    ------------------------------------------------------------------------------------
    """
    writer = pd.ExcelWriter(to_dir + "/" + excel_name + ".xlsx")
    dataframe.to_excel(writer, sheet_name='Child_pitch_shift_analysis')

    # Write Excel file
    for column in dataframe:
        # Update width of column to keep Excel sheet clean
        column_width = max(dataframe[column].astype(str).map(len).max(), len(column))
        col_idx = dataframe.columns.get_loc(column)
        writer.sheets['Child_pitch_shift_analysis'].set_column(col_idx, col_idx, column_width)

    # Save file
    writer.save()


main(sys.argv)
