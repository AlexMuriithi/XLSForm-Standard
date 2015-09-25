package com.gmail.muriithi.gibson.alex;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * XLSForm Template from Specifications defined at XLSForm org
 * 
 * @see <a href = "http://xlsform.org/"> XLSForm.org </a>
 * 
 * @author Alex Muriithi (alex.gibson.muriithi@gmail.com)
 */

public class XLSForm {

	Workbook workbook = new HSSFWorkbook();

	private static final String SURVEY = "survey";
	private static final String CHOICES = "choices";
	private static final String SETTINGS = "settings";
	private static final String INSTRUCTION = "instruction";

	public XLSForm() {
		// === Survey Sheet ===
		Sheet surveySheet = workbook.createSheet(SURVEY);
		Row surveyRow = surveySheet.createRow(0);

		Cell type = surveyRow.createCell(0);
		type.setCellValue("type");

		Cell surveyName = surveyRow.createCell(1);
		surveyName.setCellValue("name");

		Cell surveyLabel = surveyRow.createCell(2);
		surveyLabel.setCellValue("label");

		Cell surveyHint = surveyRow.createCell(3);
		surveyHint.setCellValue("hint");

		Cell constraint = surveyRow.createCell(4);
		constraint.setCellValue("constraint");

		Cell constraint_message = surveyRow.createCell(5);
		constraint_message.setCellValue("constraint_message");

		Cell required = surveyRow.createCell(6);
		required.setCellValue("required");

		Cell surveyDefault = surveyRow.createCell(7);
		surveyDefault.setCellValue("default");

		Cell relevant = surveyRow.createCell(8);
		relevant.setCellValue("relevant");

		Cell read_only = surveyRow.createCell(9);
		read_only.setCellValue("read_only");

		Cell calculation = surveyRow.createCell(10);
		calculation.setCellValue("calculation");

		Cell appearance = surveyRow.createCell(11);
		appearance.setCellValue("appearance");

		Cell hint_label_language = surveyRow.createCell(12);
		hint_label_language.setCellValue("hint/label::language");

		Cell media_image = surveyRow.createCell(13);
		media_image.setCellValue("media::image");

		Cell media_audio = surveyRow.createCell(14);
		media_audio.setCellValue("media::audio");

		Cell media_video = surveyRow.createCell(15);
		media_video.setCellValue("media::video");

		Cell media_image_language = surveyRow.createCell(16);
		media_image_language.setCellValue("media::image::language");

		Cell media_audio_language = surveyRow.createCell(17);
		media_audio_language.setCellValue("media::audio::language");

		Cell media_video_language = surveyRow.createCell(18);
		media_video_language.setCellValue("media::video::language");

		// === Choices Sheet ===
		Sheet choicesSheet = workbook.createSheet(CHOICES);
		Row choicesRow = choicesSheet.createRow(0);

		Cell list_name = choicesRow.createCell(0);
		list_name.setCellValue("list_name");

		Cell rowName = choicesRow.createCell(1);
		rowName.setCellValue("name");

		Cell rowLabel = choicesRow.createCell(2);
		rowLabel.setCellValue("label");

		Cell rowMedia = choicesRow.createCell(3);
		rowMedia.setCellValue("media");

		// === Settings Sheet ===
		Sheet settingsSheet = workbook.createSheet(SETTINGS);
		Row settingsRow = settingsSheet.createRow(0);

		Cell form_title = settingsRow.createCell(0);
		form_title.setCellValue("form_title");

		Cell form_id = settingsRow.createCell(1);
		form_id.setCellValue("form_id");

		Cell public_key = settingsRow.createCell(2);
		public_key.setCellValue("public_key");

		Cell submission_url = settingsRow.createCell(3);
		submission_url.setCellValue("submission_url");

		Cell default_language = settingsRow.createCell(4);
		default_language.setCellValue("default_language");

		Sheet instructionsSheet = workbook.createSheet(INSTRUCTION);
		Row instructionsRow = instructionsSheet.createRow(0);

		Cell se = instructionsRow.createCell(0);
		se.setCellValue("se");

		Cell instructionStandard = instructionsRow.createCell(1);
		instructionStandard.setCellValue("Standard");

		Cell instructionNotes = instructionsRow.createCell(2);
		instructionNotes.setCellValue("Notes");

		// Reading a cell value
		System.out.println("\n====\n");
		System.out
				.println("Reading cell value instructionNotes: " + instructionNotes.getRichStringCellValue().toString());
		System.out.println("\n====\n");
		try {
			FileOutputStream output = new FileOutputStream("XLSTrail.xls");
			workbook.write(output);
			workbook.close();
			System.out.println("=== File Created ===");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void main(String[] args) {
		new XLSForm();
	}
}
