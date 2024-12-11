import Foundation
import libxlsxwriter

public enum XLSXWriterError: UInt32, LocalizedError {
    case memoryMallocFailed = 1
    case creatingXlsxFile
    case creatingTempFile
    case readingTempFile
    case zipFileOperation
    case zipParameterError
    case zipBadZipFile
    case zipInternalError
    case zipFileAdd
    case zipClose
    case featureNotSupported
    case nullParameterIgnored
    case parameterValidation
    case parameterIsEmpty
    case sheetnameLengthExceeded
    case invalidSheetnameCharacter
    case sheetnameStartEndApostrophe
    case sheetnameAlreadyUsed
    case string32LengthExceeded
    case string128LengthExceeded
    case string255LengthExceeded
    case maxStringLengthExceeded
    case sharedStringIndexNotFound
    case worksheetIndexOutOfRange
    case worksheetMaxUrlLengthExceeded
    case worksheetMaxNumberUrlsExceeded
    case imageDimensions
    
    public var errorDescription: String? {
        .init(cString: lxw_strerror(.init(rawValue)))
    }
    
    static func throwIfNeeded(perform : () -> lxw_error) throws {
        let error = perform().rawValue
        guard error != 0 else { return }
        if let error = XLSXWriterError(rawValue: error) {
            throw error
        }
    }
}
