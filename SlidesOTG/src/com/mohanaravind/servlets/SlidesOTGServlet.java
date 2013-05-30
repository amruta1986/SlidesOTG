/**
 * 
 */
package com.mohanaravind.servlets;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;

/**
 * @author Aravind
 *
 */
public class SlidesOTGServlet extends HttpServlet {
	/* (non-Javadoc)
	 * @see javax.servlet.http.HttpServlet#doPost(javax.servlet.http.HttpServletRequest, javax.servlet.http.HttpServletResponse)
	 */
	@Override
	protected void doPost(HttpServletRequest req, HttpServletResponse resp)
			throws ServletException, IOException {
		// TODO Auto-generated method stub
		//super.doPost(req, resp);

		File file ;

		PrintWriter out = resp.getWriter();

		int maxFileSize = 5000 * 1024;
		int maxMemSize = 5000 * 1024;

		ServletContext context = getServletContext();
		String filePath = context.getInitParameter("file-upload");

		// Verify the content type
		String contentType = req.getContentType();
		if ((contentType.indexOf("multipart/form-data") >= 0)) {

			DiskFileItemFactory factory = new DiskFileItemFactory();
			// maximum size that will be stored in memory
			factory.setSizeThreshold(maxMemSize);
			// Location to save data that is larger than maxMemSize.
			factory.setRepository(new File("C:\\Users\\Aravind\\Desktop\\Testing\\uploads"));

			// Create a new file upload handler
			ServletFileUpload upload = new ServletFileUpload(factory);
			// maximum file size to be uploaded.
			upload.setSizeMax( maxFileSize );



			try{ 
				// Parse the request to get file items.
				List fileItems = upload.parseRequest(req);

				// Process the uploaded file items
				Iterator i = fileItems.iterator();

				out.println("<html>");
				out.println("<head>");
				out.println("<title>JSP File upload</title>");  
				out.println("</head>");
				out.println("<body>");
				while ( i.hasNext () ) 
				{
					FileItem fi = (FileItem)i.next();
					if ( !fi.isFormField () )	
					{
						// Get the uploaded file parameters
						String fieldName = fi.getFieldName();
						String fileName = fi.getName();
						boolean isInMemory = fi.isInMemory();
						long sizeInBytes = fi.getSize();
						// Write the file
						if( fileName.lastIndexOf("\\") >= 0 ){
							file = new File( filePath + 
									fileName.substring( fileName.lastIndexOf("\\"))) ;
						}else{
							file = new File( filePath + 
									fileName.substring(fileName.lastIndexOf("\\")+1)) ;
						}
						fi.write( file ) ;
						out.println("Uploaded Filename: " + filePath + 
								fileName + "<br>");
					}
				}
				out.println("</body>");
				out.println("</html>");
			}catch(Exception ex) {
				System.out.println(ex);
			}
		}else{
			out.println("<html>");
			out.println("<head>");
			out.println("<title>Servlet upload</title>");  
			out.println("</head>");
			out.println("<body>");
			out.println("<p>No file uploaded</p>"); 
			out.println("</body>");
			out.println("</html>");
		}
		
		out.close();
		out.flush();

	}

	@Override
	protected void doGet(HttpServletRequest req, HttpServletResponse resp)
			throws ServletException, IOException {
		try {
			//exportPPT();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Exports the PPT file as individual png image files
	 * @param inputFile
	 * @param outputPath
	 * @throws Exception
	 */
	private static void exportPPT(String inputFile, String outputPath) throws Exception{
		FileInputStream is = new FileInputStream(inputFile);
		SlideShow ppt = new SlideShow(is);
		is.close();

		Dimension pgsize = ppt.getPageSize();

		Slide[] slide = ppt.getSlides();
		for (int i = 0; i < slide.length; i++) {

			BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, 1);

			Graphics2D graphics = img.createGraphics();
			graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
			graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
			graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
					RenderingHints.VALUE_INTERPOLATION_BICUBIC);
			graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,
					RenderingHints.VALUE_FRACTIONALMETRICS_ON);

			graphics.setColor(Color.white);
			graphics.clearRect(0, 0, pgsize.width, pgsize.height);
			graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

			// render
			slide[i].draw(graphics);


			//Save the output
			FileOutputStream out = new FileOutputStream(outputPath + "\\" + "slide-" + (i + 1) + ".png");
			javax.imageio.ImageIO.write(img, "png", out);
			out.close();


		}
	}





}
