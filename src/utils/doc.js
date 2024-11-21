import docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZipUtils from 'jszip-utils';
import { saveAs } from 'file-saver';
import ImageModule from 'docxtemplater-image-module-free';

/**
 * 导出docx
 * @param {String} tempDocxPath 模板文件路径
 * @param {Object} data 文件中传入的数据
 * @param {String} fileName 导出文件名称
 */
export const exportDocx = async (tempDocxPath, data, fileName) => {
  // 读取模板文件的二进制内容
  JSZipUtils.getBinaryContent(tempDocxPath, async (error, content) => {
    if (error) {
      console.error('Error reading template file:', error);
      return;
    }

    // 配置图片模块选项
    const imageModuleOptions = {
      centered: false,
      getImage: getImage,
      getSize: getImageSize
    };

    const imageModule = new ImageModule(imageModuleOptions);

    // 创建PizZip实例，并加载模板内容
    const zip = new PizZip(content);

    // 创建 docxtemplater 实例并传递模板和模块
    const doc = new docxtemplater(zip, { modules: [imageModule] });

    try {
      // 使用 renderAsync 解析数据并渲染文档
      await doc.renderAsync(data);  // 异步渲染文档
      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      // 触发文件下载
      saveAs(out, fileName);
    } catch (err) {
      console.error('Error during docxtemplater processing:', err);
    }
  });
};

/**
 * 获取图片内容
 * @param {String} tagValue 图片路径
 * @returns {Promise<Buffer>} 图片二进制数据
 */
const getImage = (tagValue) => {
  return new Promise((resolve, reject) => {
    JSZipUtils.getBinaryContent(tagValue, (error, content) => {
      if (error) {
        console.error(`Error loading image from ${tagValue}, using fallback image.`);
        // 如果图片加载失败，使用备用图片
        JSZipUtils.getBinaryContent('/failImg.png', (fallbackError, fallbackContent) => {
          if (fallbackError) {
            reject(`Error loading fallback image: ${fallbackError}`);
          } else {
            resolve(fallbackContent); // 返回备用图片
          }
        });
      } else {
        resolve(content); // 返回原始图片
      }
    });
  });
};

/**
 * 获取图片尺寸（最大宽度为300px）
 * @param {HTMLImageElement} img 图片元素
 * @param {String} tagValue 图片路径
 * @param {String} tagName 标签名
 * @returns {Promise<[number, number]>} 返回图片的宽度和高度
 */
const getImageSize = (img, tagValue) => {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.src = tagValue;

    image.onload = () => {
      let imgWidth = image.width;
      let imgHeight = image.height;

      if (imgWidth > 300) {
        const scale = 300 / imgWidth;
        imgWidth = 300;
        imgHeight = imgHeight * scale;
      }

      resolve([imgWidth, imgHeight]);
    };

    image.onerror = () => {
      console.error(`Error loading image size for ${tagValue}, using fallback image.`);
      // 在图片加载失败时使用备用图片进行尺寸计算
      image.src = '/failImg.png';

      image.onload = () => {
        let fallbackWidth = image.width;
        let fallbackHeight = image.height;

        if (fallbackWidth > 300) {
          const scale = 300 / fallbackWidth;
          fallbackWidth = 300;
          fallbackHeight = fallbackHeight * scale;
        }

        resolve([fallbackWidth, fallbackHeight]);
      };

      image.onerror = (e) => {
        console.error(`Error loading fallback image size: ${e}`);
        reject(e);
      };
    };
  });
};
