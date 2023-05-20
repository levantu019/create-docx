import * as docx from "docx";

const dateDocument = "Hà Nội, ngày 09 tháng 10 năm 2023";
const p_organDocument = "BỘ QUỐC PHÒNG";
const organDocument = "BINH CHỦNG HOÁ HỌC";
const nameDocument = "BÁO CÁO";

const styleDocument = {
  paragraphStyles: [
    {
      id: "paragraph-style-1",
      name: "Paragrap Style 1",
      basedOn: "Normal",
      next: "Normal",
      run: {
        color: "000000",
        bold: true,
        size: 28
      },
      paragraph: {
        alignment: docx.AlignmentType.CENTER
      }
    },
    {
      id: "paragraph-style-main",
      name: "Paragraph Style Main",
      basedOn: "Normal",
      quickFormat: true,
      run: {
        color: "000000",
        size: 28
      },
      paragraph: {
        spacing: {
          line: 276,
          before: 20 * 72 * 0.1,
          after: 20 * 72 * 0.05
        }
      }
    },
    {
      id: "title-table",
      name: "Titlw in Table",
      basedOn: "Normal",
      quickFormat: true,
      run: {
        color: "000000",
        size: 28,
        bold: true
      },
      paragraph: {
        alignment: docx.AlignmentType.CENTER
      }
    }
  ]
};

class sections {
  static headerDocument() {
    const table = new docx.Table({
      width: {
        size: "18cm",
        type: docx.WidthType.AUTO
      },
      indent: {
        size: "-2cm"
      },
      borders: {
        top: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
        insideHorizontal: {
          style: docx.BorderStyle.NONE,
          size: 0,
          color: "FFFFFF"
        },
        insideVertical: {
          style: docx.BorderStyle.NONE,
          size: 0,
          color: "FFFFFF"
        }
      },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [
                new docx.Paragraph({
                  style: "paragraph-style-1",
                  children: [
                    new docx.TextRun({
                      text: p_organDocument,
                      bold: false
                    })
                  ]
                }),
                new docx.Paragraph({
                  style: "paragraph-style-1",
                  spacing: {
                    line: 276
                  },
                  children: [
                    new docx.TextRun({
                      text: organDocument
                    })
                  ]
                })
              ]
            }),
            new docx.TableCell({
              children: [
                new docx.Paragraph({
                  style: "paragraph-style-1",
                  children: [
                    new docx.TextRun({
                      text: "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM"
                    })
                  ]
                }),
                new docx.Paragraph({
                  style: "paragraph-style-1",
                  spacing: {
                    line: 276
                  },
                  children: [
                    new docx.TextRun({
                      text: "Độc lập - Tự do - Hạnh phúc"
                    })
                  ]
                })
              ]
            })
          ]
        })
      ]
    });

    return table;
  }
}

// Create file get all equipments (published, parent)
export const DocumentAllEquipments = (data) => {
  const document = new docx.Document({
    styles: styleDocument,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: docx.convertMillimetersToTwip(25),
              right: docx.convertMillimetersToTwip(15),
              bottom: docx.convertMillimetersToTwip(20),
              left: docx.convertMillimetersToTwip(35)
            }
          }
        },
        children: [
          sections.headerDocument(),

          new docx.Paragraph({
            style: "paragraph-style-main",
            alignment: docx.AlignmentType.RIGHT,
            children: [
              new docx.TextRun({
                text: dateDocument,
                italics: true
              })
            ]
          }),

          new docx.Paragraph({
            style: "paragraph-style-1",
            spacing: {
              line: 276
            },
            children: [
              new docx.TextRun({
                text: nameDocument
              })
            ]
          }),

          //
          new docx.Paragraph({
            children: []
          }),

          //
          new docx.Table({
            width: {
              size: "16cm",
              type: docx.WidthType.AUTO
            },
            rows: [
              new docx.TableRow({
                children: [
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "STT"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Tên"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Tình trạng"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Số hiệu"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Phân cấp"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Đơn vị"
                          })
                        ]
                      })
                    ]
                  })
                ]
              }),
              ...data.map(
                (item, index) =>
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: (index + 1).toString(),
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.name,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.tinhtrang,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.sohieu,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.phancap,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.donvi,
                                bold: false
                              })
                            ]
                          })
                        ]
                      })
                    ]
                  })
              )
            ]
          })
        ]
      }
    ]
  });

  return document;
};

// Create file get childs equipments
export const DocumentChildEquipments = (data) => {
  const document = new docx.Document({
    styles: styleDocument,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: docx.convertMillimetersToTwip(25),
              right: docx.convertMillimetersToTwip(15),
              bottom: docx.convertMillimetersToTwip(20),
              left: docx.convertMillimetersToTwip(35)
            }
          }
        },
        children: [
          sections.headerDocument(),

          new docx.Paragraph({
            style: "paragraph-style-main",
            alignment: docx.AlignmentType.RIGHT,
            children: [
              new docx.TextRun({
                text: dateDocument,
                italics: true
              })
            ]
          }),

          new docx.Paragraph({
            style: "paragraph-style-1",
            spacing: {
              line: 276
            },
            children: [
              new docx.TextRun({
                text: nameDocument
              })
            ]
          }),

          //
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Tên: " + data.parent.name
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Số hiệu: " + data.parent.sohieu
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Đơn vị: " + data.parent.donvi
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Loại: " + data.parent.loai
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Xuất xứ: " + data.parent.xuatxu
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Năm sản xuất: " + data.parent.namsx
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Tình trạng: " + data.parent.tinhtrang
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Phân cấp: " + data.parent.phancap
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Chức năng: " + data.parent.chucnang
          }),
          new docx.Paragraph({
            style: "paragraph-style-main",
            text: "- Danh sách"
          }),

          //
          new docx.Table({
            width: {
              size: "16cm",
              type: docx.WidthType.AUTO
            },
            rows: [
              new docx.TableRow({
                children: [
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "STT"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Tên"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Tình trạng"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Số hiệu"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Phân cấp"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Đơn vị"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Loại"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Xuất xứ"
                          })
                        ]
                      })
                    ]
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        style: "title-table",
                        children: [
                          new docx.TextRun({
                            text: "Năm sản xuất"
                          })
                        ]
                      })
                    ]
                  })
                ]
              }),
              ...data.childs.map(
                (item, index) =>
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: (index + 1).toString(),
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.name,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.tinhtrang,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.sohieu,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.phancap,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.donvi,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.loai,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.xuatxu,
                                bold: false
                              })
                            ]
                          })
                        ]
                      }),
                      new docx.TableCell({
                        children: [
                          new docx.Paragraph({
                            style: "title-table",
                            children: [
                              new docx.TextRun({
                                text: item.namsx,
                                bold: false
                              })
                            ]
                          })
                        ]
                      })
                    ]
                  })
              )
            ]
          })
        ]
      }
    ]
  });

  return document;
};
