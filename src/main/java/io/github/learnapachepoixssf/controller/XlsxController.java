package io.github.learnapachepoixssf.controller;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.util.StopWatch;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import io.github.learnapachepoixssf.service.XlsxService;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Controller
public class XlsxController {

	static final String XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	@Autowired
	protected XlsxService xlsxService;

	@GetMapping("/test-widgets")
	public void getTestWidgets(@RequestParam(value = "rows", defaultValue = "10", required = false) int rows,
			HttpServletResponse response) throws IOException {
		StopWatch stopWatch = new StopWatch("testdata");
		stopWatch.start();
		response.setContentType(XLSX_CONTENT_TYPE);
		response.setHeader("Content-Disposition", "attachment; filename=\"test-widgets.xlsx\"");
		xlsxService.writeOutTestWidgets(response.getOutputStream(), rows);
		stopWatch.stop();
		log.info("Time to write out test widgets: " + stopWatch.toString());
	}

	@GetMapping("/download-widgets")
	@ResponseBody
	public void downloadWidgets(HttpServletResponse response) throws IOException {
		StopWatch stopWatch = new StopWatch("download");
		stopWatch.start();
		response.setContentType(XLSX_CONTENT_TYPE);
		response.setHeader("Content-Disposition", "attachment; filename=\"widgets.xlsx\"");
		int rowCount = xlsxService.writeOutSavedWidgets(response.getOutputStream());
		stopWatch.stop();
		log.info("Time to write out {} saved widgets: {}", rowCount, stopWatch.toString());
	}

	@PostMapping("/truncate-widgets")
	public String truncateWidgets() {
		StopWatch stopWatch = new StopWatch("truncate");
		stopWatch.start();
		xlsxService.truncateWidgets();
		stopWatch.stop();
		log.info("Time to truncat widgets: {}", stopWatch.toString());
		return "redirect:/index.html?truncated=true";
	}

	@PostMapping("/upload-widgets")
	public String uploadWidgets(@RequestParam("file") MultipartFile file,
			@RequestParam(value = "batchSize", defaultValue = "1000") int batchSize,
			@RequestParam(value = "parseType", defaultValue = "sax") String parseType,
			@RequestParam(value = "persistenceType", defaultValue = "jdbc") String persistenceType,
			RedirectAttributes redirectAttributes) throws Exception {
		Path savedFile = null;
		try {
			StopWatch stopWatch = new StopWatch(parseType + "-" + persistenceType);
			stopWatch.start();
			savedFile = save(file);
			log.info("Uploaded file: " + savedFile);
			int rowCount = xlsxService.saveWidgets(savedFile, batchSize, parseType, persistenceType);
			stopWatch.stop();
			log.info("Time to save {} widgets: {}", rowCount, stopWatch.toString());
			return "redirect:/index.html?uploaded=true";
		} finally {
			if (savedFile != null) {
				FileUtils.deleteQuietly(savedFile.toFile());
			}
		}
	}

	protected Path save(MultipartFile file) throws IOException {
		String filename = StringUtils.cleanPath(file.getOriginalFilename());
		if (file.isEmpty()) {
			throw new IllegalArgumentException("Cannot process empty file: " + filename);
		}
		if (filename.contains("..")) {
			// This is a security check
			throw new IllegalArgumentException("Cannot store file with relative path: " + filename);
		}
		try (InputStream inputStream = file.getInputStream()) {
			Path tempLocation = Files.createTempDirectory("widget-files");
			Path tempFile = tempLocation.resolve(filename);
			Files.copy(inputStream, tempFile, StandardCopyOption.REPLACE_EXISTING);
			return tempFile;
		}
	}
}
