package org.one2team.highcharts.server.export;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.awt.image.DataBufferByte;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hslf.model.Picture;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.one2team.highcharts.server.export.util.SVGRendererInternal;

import javax.imageio.ImageIO;

public class HighchartsExporter<T> {

	public HighchartsExporter(ExportType type, SVGRendererInternal<T> internalRenderer) {
		this.type = type;
		this.renderer = 
			new SVGStreamRenderer<T> (new SVGRenderer<T> (internalRenderer),
				                     type.getTranscoder ());
	}




	public void export (T chartOptions,
			                T globalOptions,
			                File file) {
		
		OutputStream fos = null;
		try {
            if (this.type == ExportType.ppt){
                File tempFile = new File(file.getAbsolutePath().replace(".ppt",".png")) ;
                fos = render (chartOptions, globalOptions, tempFile);

                SlideShow slideShow = new SlideShow();
                Slide slide = slideShow.createSlide();


                BufferedImage image = ImageIO.read(tempFile);
                Dimension pageSize =  new Dimension();
                pageSize.setSize(image.getWidth(), image.getHeight());
                slideShow.setPageSize(pageSize );
                Picture picture = new Picture(slideShow.addPicture(tempFile, Picture.PNG));
                picture.setAnchor(new Rectangle(0, 0, image.getWidth(), image.getHeight()));
                slide.addShape(picture);

                slideShow.write(new FileOutputStream(file));

            } else
			    fos = render (chartOptions, globalOptions, file);

		} catch (Exception e) {
			e.printStackTrace ();
			throw (new RuntimeException (e));
		} finally {
			if (fos != null)
				IOUtils.closeQuietly (fos);
		}
	}

	private OutputStream render (T chartOptions,
			                         T globalOptions,
                               File file) throws FileNotFoundException {
		FileOutputStream fos;
		renderer.setChartOptions (chartOptions)
				    .setGlobalOptions (globalOptions)
				    .setOutputStream (fos = new FileOutputStream (file))
				    .render ();
		return fos;
	}

	public SVGStreamRenderer<T> getRenderer () {
		return renderer;
	}

	public ExportType getType () {
		return type;
	}

	private final SVGStreamRenderer<T> renderer;

	private final ExportType type;
}
